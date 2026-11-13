(() => {
  'use strict';
  const records = new WeakMap(), active = new Set();
  function clear(image) {
    records.delete(image); active.delete(image);
    image.removeAttribute('src'); image.hidden = true; image.dataset.previewState = 'empty';
  }
  async function show(image, data, current) {
    if (typeof current !== 'function' || !current()) return false;
    clear(image);
    if (typeof data !== 'string' || data.length > 300000 || !/^data:image\/(?:jpeg|png);base64,[A-Za-z0-9+/=]+$/.test(data) || typeof current !== 'function' || !current()) return false;
    const record = {data, current}; records.set(image, record); active.add(image);
    image.dataset.previewState = 'decoding'; image.src = data;
    try {
      await image.decode();
      await new Promise(requestAnimationFrame);
      if (records.get(image) !== record || !current()) { if(records.get(image) === record) clear(image); return false; }
      if (!image.naturalWidth || !image.naturalHeight) throw new Error();
      image.hidden = false; image.dataset.previewState = 'ready'; return true;
    } catch (_) {
      if (records.get(image) === record) { image.hidden = true; image.dataset.previewState = 'failed'; }
      return false;
    }
  }
  async function retry(image) {
    const record = records.get(image);
    if (!record || !record.current()) return false;
    const data = record.data, current = record.current;
    clear(image); const pending={data,current};records.set(image,pending);active.add(image);
    await new Promise(requestAnimationFrame);
    if (!current() || records.get(image)!==pending) {if(records.get(image)===pending)clear(image);return false;}
    return show(image, data, current);
  }
  window.OfficePdfPreview = {show, retry, clear};
  for (const event of ['officeauth:unauthenticated', 'officeauth:authenticated', 'pagehide']) window.addEventListener(event, () => { for (const image of [...active]) clear(image); });
})();
