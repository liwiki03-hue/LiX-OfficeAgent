/* Display-only adapter. Not loaded by production UI until host controls exist.
 * Call only with host-validated source; return SVG MUST pass Python validate_svg.
 * Never inserts returned SVG into the parent, exports files, or calls an API. */
(() => {
  'use strict';
  const base = new URL('.', document.currentScript?.src || document.baseURI);
  const worker = new URL('diagram_worker.js', base);
  const vendor = new URL('vendor/mermaid.min.js', base);
  const local = base.protocol === 'file:' || (base.protocol === 'http:' && ['127.0.0.1', 'localhost', '[::1]'].includes(base.hostname));
  const errors = new Set(['renderer_failed', 'unsafe_svg', 'render_timeout', 'render_cancelled', 'unsupported_source']);
  const details = new Set(['shape','foreignObject','symbol','use','filter','clipPath']);
  const fail = (code, detail) => Object.assign(new Error('图表未能安全显示，请简化后重试'), {code, ...(details.has(detail) ? {detail} : {})});
  let active = 0;
  function render(source, options = {}) {
    return new Promise((resolve, reject) => {
      if (!options || typeof options !== 'object' || Array.isArray(options)) return reject(fail('unsupported_source'));
      const {signal, timeoutMs = 8000} = options;
      if (signal !== undefined && !(signal instanceof AbortSignal)) return reject(fail('unsupported_source'));
      if (active >= 2) return reject(fail('renderer_failed'));
      if (!local || typeof source !== 'string' || !source.length || source.length > 8192 || /[^\x09\x0a\x0d\x20-\x7e]/.test(source)) return reject(fail('unsupported_source'));
      if (!Number.isInteger(timeoutMs) || timeoutMs < 100 || timeoutMs > 8000) return reject(fail('unsupported_source'));
      if (signal?.aborted) return reject(fail('render_cancelled'));
      const token = Array.from(crypto.getRandomValues(new Uint8Array(16)), n => n.toString(16).padStart(2, '0')).join('');
      const frame = document.createElement('iframe');
      frame.setAttribute('sandbox', 'allow-scripts');
      frame.setAttribute('referrerpolicy', 'no-referrer');
      frame.setAttribute('aria-hidden', 'true');
      frame.style.cssText = 'position:fixed;left:-10000px;top:0;width:1200px;height:1000px;border:0;';
      active += 1;
      let settled = false, started = false;
      const finish = (error, result, detail) => {
        if (settled) return;
        settled = true; active -= 1; clearTimeout(timer); window.removeEventListener('message', receive);
        signal?.removeEventListener('abort', cancel); frame.remove();
        if (error) reject(fail(error, detail)); else resolve(result);
      };
      const cancel = () => finish('render_cancelled');
      const receive = event => {
        if (event.source !== frame.contentWindow || event.origin !== 'null') return;
        const data = event.data;
        if (!data || data.channel !== 'office-diagram-v1') return;
        if (data.type === 'ready' && !started) {
          if (data.bridgeAbsent !== true) return finish('renderer_failed');
          started = true;
          frame.contentWindow.postMessage({channel:'office-diagram-v1', type:'render', token, source}, '*');
        } else if (started && data.token === token) {
          if (data.type === 'error') finish(errors.has(data.code) ? data.code : 'renderer_failed', null, data.detail);
          else if (data.type === 'result') {
            if (typeof data.svg !== 'string' || data.svg.length > 1048576 || data.bridgeAbsent !== true) return finish('unsafe_svg');
            finish(null, {svg:data.svg, bridgeAbsent:true});
          }
        }
      };
      const timer = setTimeout(() => finish('render_timeout'), timeoutMs);
      window.addEventListener('message', receive); signal?.addEventListener('abort', cancel, {once:true});
      // Only fixed local URLs enter srcdoc; source is sent later as structured data.
      const escape = value => String(value).replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
      const csp = `default-src 'none'; script-src ${vendor.href} ${worker.href}; style-src 'unsafe-inline'; connect-src 'none'; img-src 'none'; font-src 'none'; object-src 'none'; frame-src 'none'; base-uri 'none'; form-action 'none'`;
      frame.srcdoc = `<!doctype html><meta http-equiv="Content-Security-Policy" content="${escape(csp)}"><script src="${escape(vendor.href)}"></script><script src="${escape(worker.href)}"></script>`;
      try { document.body.append(frame); } catch (_) { finish('renderer_failed'); }
    });
  }
  window.OfficeDiagramRenderer = Object.freeze({render});
})();
