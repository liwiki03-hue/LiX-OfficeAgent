/* Runs only inside opaque-origin sandbox. No export, host API, or network. */
(() => {
  'use strict';
  const NS = 'http://www.w3.org/2000/svg';
  const allowed = new Set(['svg','g','path','rect','circle','ellipse','line','polyline','polygon','text','tspan','defs','marker']);
  const geometry = new Set(['id','class','x','y','x1','y1','x2','y2','cx','cy','r','rx','ry','width','height','refX','refY','markerWidth','markerHeight','markerUnits','orient','viewBox','d','points','transform','dx','dy','textLength','preserveAspectRatio']);
  const properties = ['fill','stroke','stroke-width','stroke-dasharray','stroke-linecap','stroke-linejoin','fill-rule','opacity','fill-opacity','stroke-opacity','text-anchor','dominant-baseline','alignment-baseline','font-size','font-style'];
  const bridgeAbsent = () => {
    if (!(typeof window.pywebview === 'undefined') || window === parent) return false;
    try { void parent.document; return false; } catch (_) { return true; }
  };
  const send = payload => parent.postMessage({channel:'office-diagram-v1', ...payload}, '*');
  const bad = (detail = 'shape') => { throw Object.assign(new Error('unsafe_svg'), {detail}); };
  function flatten(root) {
    const view = root.getAttribute('viewBox')?.trim().split(/[ ,]+/).map(Number);
    if (!view || view.length !== 4 || !view.every(Number.isFinite) || view[2] < 1 || view[3] < 1 || view[2] > 4096 || view[3] > 4096 || view[2]*view[3] > 8388608) bad();
    // Sequence renderer predeclares unused actor icons. Restricted participants
    // never use them; reject use elements, discard only inert defs symbols.
    if (root.querySelector('use')) bad('use');
    let count = 0, text = 0;
    function clone(el, depth = 0) {
      if (++count > 4096 || depth > 32 || el.namespaceURI !== NS) bad();
      if (['style','title','desc'].includes(el.localName)) return null; // non-business metadata/CSS only
      if (el.localName === 'symbol' && el.closest('defs')) return null;
      if (!allowed.has(el.localName) || (depth && el.localName === 'svg')) bad(el.localName);
      const out = document.createElementNS(NS, el.localName), css = getComputedStyle(el);
      const unit = value => /^[-+]?(?:\d*\.\d+|\d+)em$/.test(value) ? String(parseFloat(value)*parseFloat(css.fontSize)) : value.replace(/px\b/g, '');
      for (const attr of el.attributes) {
        if (geometry.has(attr.name)) {
          let value = unit(attr.value);
          // Mermaid class markers reserve a huge unclipped viewport without a
          // viewBox. Bound that viewport, not the marker's coordinate geometry.
          if (el.localName === 'marker' && ['markerWidth','markerHeight'].includes(attr.name) && !el.hasAttribute('viewBox')) value = String(Math.min(64, Number(value)));
          out.setAttribute(attr.name, value);
        }
      }
      for (const name of properties) {
        let value = css.getPropertyValue(name).trim();
        if (!value || (name === 'stroke-dasharray' && value === 'none')) continue;
        if (/url\(/i.test(value)) bad();
        if (/^rgba\(\s*0\s*,\s*0\s*,\s*0\s*,\s*0\s*\)$/.test(value)) value = 'none';
        out.setAttribute(name, unit(value));
      }
      // Fixed default theme: ER wrappers can inherit the shape fill.
      // Flatten labels with explicit contrasting passive colors instead.
      if (['text','tspan'].includes(el.localName)) {
        out.setAttribute('fill', '#222'); out.setAttribute('stroke', 'none');
      }
      const labelBackdrop = el.localName === 'rect' && el.closest('.edgeLabel');
      if (labelBackdrop) { out.setAttribute('fill', '#fff'); out.setAttribute('stroke', 'none'); }
      out.setAttribute('font-family', 'Arial');
      out.setAttribute('font-weight', Number(css.fontWeight) >= 600 ? '700' : '400');
      for (const name of ['marker-start','marker-mid','marker-end']) {
        const value = css.getPropertyValue(name).trim();
        if (!value || value === 'none') continue;
        const match = /^url\(["']?([^"')]+)["']?\)$/.exec(value);
        if (!match) bad();
        const target = match[1], split = target.lastIndexOf('#');
        if (split < 0 || (split > 0 && target.slice(0,split) !== document.baseURI.split('#')[0])) bad();
        const id = target.slice(split+1);
        if (!/^[A-Za-z_][A-Za-z0-9_-]{0,127}$/.test(id)) bad();
        out.setAttribute(name, `url(#${id})`);
      }
      for (const child of el.childNodes) {
        if (child.nodeType === 1) { const nested = clone(child, depth+1); if (nested) out.append(nested); }
        else if (child.nodeType === 3) { text += child.textContent.length; if (text > 8192) bad(); out.append(document.createTextNode(child.textContent)); }
      }
      return out;
    }
    const output = clone(root);
    output.setAttribute('viewBox', view.join(' ')); output.setAttribute('width', String(view[2])); output.setAttribute('height', String(view[3]));
    const svg = new XMLSerializer().serializeToString(output);
    if (new TextEncoder().encode(svg).length > 1048576) bad();
    return svg;
  }
  let used = false;
  window.addEventListener('message', async event => {
    if (event.source !== parent || used) return;
    const data = event.data;
    if (!data || data.channel !== 'office-diagram-v1' || data.type !== 'render' || !/^[0-9a-f]{32}$/.test(data.token)) return;
    used = true;
    if (!bridgeAbsent() || typeof data.source !== 'string' || !data.source.length || data.source.length > 8192 || /[^\x09\x0a\x0d\x20-\x7e]/.test(data.source)) return send({type:'error', token:data.token, code:'renderer_failed'});
    const stage = document.createElement('div'); document.body.append(stage);
    try {
      mermaid.initialize({startOnLoad:false, securityLevel:'strict', suppressErrorRendering:true, maxTextSize:8192, maxEdges:36, deterministicIds:true, deterministicIDSeed:data.token, theme:'default', fontFamily:'Arial', htmlLabels:false, flowchart:{htmlLabels:false}, class:{htmlLabels:false}, er:{useMaxWidth:false}});
      const result = await mermaid.render('d'+data.token, data.source, stage);
      // Untrusted render output stays INSIDE the isolated frame for computed style.
      if (typeof result.svg !== 'string' || new TextEncoder().encode(result.svg).length > 1048576) bad();
      stage.innerHTML = result.svg;
      const root = stage.querySelector('svg'); if (!root || !bridgeAbsent()) bad();
      const svg = flatten(root);
      if (!bridgeAbsent()) bad();
      send({type:'result', token:data.token, svg, bridgeAbsent:true});
    } catch (error) {
      send({type:'error', token:data.token, code:error?.message === 'unsafe_svg' ? 'unsafe_svg' : 'renderer_failed', detail:error?.detail});
    } finally { stage.remove(); }
  });
  send({type:'ready', bridgeAbsent:bridgeAbsent()});
})();
