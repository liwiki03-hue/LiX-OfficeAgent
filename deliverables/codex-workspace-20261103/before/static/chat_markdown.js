"use strict";
(() => {
  function inline(parent, text) {
    const pattern = /(!?\[([^\]\n]*)\]\(([^\s)]+)\)|`([^`\n]+)`|\*\*([^*\n]+)\*\*|\*([^*\n]+)\*)/g;
    let offset = 0, match;
    while ((match = pattern.exec(text))) {
      parent.append(document.createTextNode(text.slice(offset,match.index)));
      let node;
      if (match[4]) { node = document.createElement('code'); node.textContent = match[4]; }
      else if (match[5] || match[6]) { node = document.createElement(match[5] ? 'strong' : 'em'); node.textContent = match[5] || match[6]; }
      else if (!match[1].startsWith('!') && /^https?:\/\//i.test(match[3])) {
        try {
          const url = new URL(match[3]);
          if (url.username || url.password) throw new Error();
          node = document.createElement('a'); node.href = url.href; node.target = '_blank'; node.rel = 'noopener noreferrer'; node.textContent = match[2];
        } catch (_) { node = document.createTextNode(match[1]); }
      } else node = document.createTextNode(match[1]);
      parent.append(node); offset = pattern.lastIndex;
    }
    parent.append(document.createTextNode(text.slice(offset)));
  }
  function render(parent, raw) {
    const fragment = document.createDocumentFragment(), lines = raw.split(/\r?\n/);
    const add = (tag,text,where=fragment) => { const node = document.createElement(tag); inline(node,text); where.append(node); return node; };
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      if (/^\s*```/.test(line)) {
        const block = document.createElement('div'); block.className = 'chat-code';
        const language = line.replace(/^\s*```/,'').trim();
        const code = [], copy = document.createElement('button');
        while (++i < lines.length && !/^\s*```/.test(lines[i])) code.push(lines[i]);
        const text = code.join('\n'); copy.type = 'button'; copy.className = 'btn secondary'; copy.textContent = '复制代码' + (language ? ' · ' + language : '');
        copy.onclick = async () => { try { await navigator.clipboard.writeText(text); copy.textContent = '代码已复制'; } catch (_) { copy.textContent = '剪贴板不可用，请选择代码复制'; } };
        const pre = document.createElement('pre'), content = document.createElement('code'); content.textContent = text; pre.append(content); block.append(copy,pre); fragment.append(block); continue;
      }
      if (!line.trim()) continue;
      const heading = /^(#{1,6})\s+(.*)$/.exec(line);
      if (heading) { add('h' + heading[1].length, heading[2]); continue; }
      if (i+1 < lines.length && line.includes('|') && /^\s*\|?\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)+\|?\s*$/.test(lines[i+1])) {
        const wrapper = document.createElement('div'); wrapper.className = 'chat-table'; wrapper.tabIndex = 0;
        const table = document.createElement('table'); wrapper.append(table); fragment.append(wrapper);
        const cells = value => value.trim().replace(/^\||\|$/g,'').split('|');
        const tr = document.createElement('tr'); for (const cell of cells(line)) add('th',cell.trim(),tr); table.append(tr); i++;
        while (i+1 < lines.length && lines[i+1].includes('|') && lines[i+1].trim()) { const row = document.createElement('tr'); for (const cell of cells(lines[++i])) add('td',cell.trim(),row); table.append(row); }
        continue;
      }
      if (/^\s*>/.test(line)) { add('blockquote',line.replace(/^\s*>\s?/,'')); continue; }
      const item = /^\s*(?:([-+*])|\d+[.)])\s+(.*)$/.exec(line);
      if (item) {
        const ordered = !item[1], list = add(ordered ? 'ol' : 'ul','');
        if (ordered) list.start = parseInt(line.trim(),10);
        add('li',item[2],list);
        while (i+1 < lines.length) {
          const next = /^\s*(?:([-+*])|\d+[.)])\s+(.*)$/.exec(lines[i+1]);
          if (!next || Boolean(next[1]) === ordered) break;
          add('li',next[2],list); i++;
        }
        continue;
      }
      add('p',line);
    }
    parent.replaceChildren(fragment);
  }
  window.OfficeMarkdown = Object.freeze({render});
})();
