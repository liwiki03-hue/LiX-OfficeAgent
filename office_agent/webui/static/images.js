/* Read-only raster presentation. No HTML, network URLs, downloads or write permissions. */
"use strict";
(() => {
  const MAX_BYTES = 10_000_000, MAX_IMAGES = 4, MAX_PIXELS = 40_000_000;
  const MAX_ENCODED = Math.ceil(MAX_BYTES / 3) * 4;
  const text = value => typeof value === "string" ? value.slice(0, 1000) : "";
  function dimensions(bytes, mime) {
    const u16 = i => (bytes.charCodeAt(i) << 8) | bytes.charCodeAt(i + 1);
    const u32 = i => u16(i) * 65536 + u16(i + 2);
    if (mime === "image/png") {
      if (bytes.length < 33 || bytes.slice(0, 8) !== "\x89PNG\r\n\x1a\n" || u32(8) !== 13 || bytes.slice(12, 16) !== "IHDR") return null;
      return [u32(16), u32(20)];
    }
    if (bytes.slice(0, 2) !== "\xff\xd8") return null;
    // Inspect JPEG segment lengths before handing anything to the browser decoder.
    let offset = 2;
    while (offset + 4 <= bytes.length) {
      if (bytes.charCodeAt(offset++) !== 255) return null;
      while (offset < bytes.length && bytes.charCodeAt(offset) === 255) offset++;
      const marker = bytes.charCodeAt(offset++);
      if (marker === 217 || marker === 218 || marker === 0) return null;
      if (marker === 216 || marker === 1 || (marker >= 208 && marker <= 215)) continue;
      const length = u16(offset);
      if (length < 2 || offset + length > bytes.length) return null;
      if ([192, 193, 194, 195, 197, 198, 199, 201, 202, 203, 205, 206, 207].includes(marker)) {
        return length >= 8 ? [u16(offset + 5), u16(offset + 3)] : null;
      }
      offset += length;
    }
    return null;
  }
  function validateImage(value) {
    if (!value || typeof value !== "object" || typeof value.data_url !== "string") return null;
    const url = value.data_url;
    if (url.length > MAX_ENCODED + 40) return null;
    const match = /^data:(image\/(?:png|jpeg));base64,([A-Za-z0-9+/]+={0,2})$/.exec(url);
    if (!match || match[2].length % 4 !== 0) return null;
    let bytes;
    try { bytes = atob(match[2]); if (btoa(bytes) !== match[2]) return null; } catch (_) { return null; }
    if (!bytes.length || bytes.length > MAX_BYTES) return null;
    const size = dimensions(bytes, match[1]);
    if (!size || !size.every(n => Number.isInteger(n) && n > 0 && n <= 32768) || size[0] * size[1] > MAX_PIXELS) return null;
    return {data_url: url, alt: text(value.alt) || "图片预览", caption: text(value.caption), width: size[0], height: size[1]};
  }
  function showLarge(image, trigger) {
    const dialog = document.createElement("dialog"); dialog.className = "image-lightbox";
    dialog.setAttribute("aria-label", "放大查看图片");
    const close = document.createElement("button"); close.type = "button"; close.className = "btn secondary"; close.textContent = "关闭放大图片";
    const img = document.createElement("img"); img.alt = image.alt; img.src = image.data_url;
    const caption = document.createElement("p"); caption.textContent = image.caption || image.alt;
    close.onclick = () => dialog.close();
    dialog.addEventListener("keydown", event => {
      event.stopPropagation(); // Do not submit a task or navigate history underneath this modal.
      if (event.key === "Escape") { event.preventDefault(); dialog.close(); }
    });
    dialog.addEventListener("click", event => { if (event.target === dialog) dialog.close(); });
    dialog.addEventListener("close", () => { dialog.remove(); if (trigger.isConnected) trigger.focus(); }, {once: true});
    dialog.append(close, img, caption); document.body.append(dialog); dialog.showModal(); close.focus();
  }
  function appendImages(host, values, options = {}) {
    if (!Array.isArray(values) || !values.length) return;
    const group = document.createElement("section"); group.className = "image-results"; group.setAttribute("aria-label", "只读图片结果");
    values.slice(0, MAX_IMAGES).forEach(value => {
      const image = validateImage(value), card = document.createElement("figure"); card.className = "image-card";
      const caption = document.createElement("figcaption"); caption.textContent = text(value?.caption) || text(value?.alt) || "图片预览";
      if (!image) {
        const warning = document.createElement("p"); warning.className = "image-warning";
        warning.textContent = "图片未显示：仅接受有界的 PNG/JPEG data URL；外部地址、SVG、无效内容或超限图片均拒绝。";
        card.append(warning, caption); group.append(card); return;
      }
      const enlarge = document.createElement("button"); enlarge.type = "button"; enlarge.className = "image-enlarge";
      enlarge.setAttribute("aria-label", "放大查看：" + image.alt);
      const img = document.createElement("img"); img.alt = image.alt; img.loading = "lazy"; img.decoding = "async";
      img.onerror = () => { enlarge.disabled = true; img.removeAttribute("src"); caption.textContent = "图片解码失败；未显示或保存内容。"; };
      img.onload = () => {
        if (img.naturalWidth !== image.width || img.naturalHeight !== image.height) {
          enlarge.disabled = true; img.removeAttribute("src"); caption.textContent = "图片尺寸校验失败。";
        }
      };
      img.src = image.data_url; enlarge.append(img); enlarge.onclick = () => showLarge(image, enlarge);
      card.append(enlarge, caption);
      const note = document.createElement("small");
      if (options.transactionalSave === true) {
        note.textContent = "仅供预览；保存请使用结果卡下方的独立确认入口。";
      } else {
        const save = document.createElement("button"); save.type = "button"; save.className = "btn secondary image-save";
        save.textContent = "保存到本地（未启用）"; save.disabled = true;
        save.title = "此只读预览不提供保存权限；不能绕过授权、独立确认与审计。";
        card.append(save);
        note.textContent = "仅供预览 · 此预览不提供保存操作";
      }
      card.append(note); group.append(card);
    });
    if (values.length > MAX_IMAGES) {
      const limit = document.createElement("p"); limit.textContent = "本次最多展示 4 张图片；其余未展示。"; group.append(limit);
    }
    host.append(group);
  }
  window.addEventListener("officeauth:unauthenticated", () => {
    document.querySelectorAll(".image-lightbox").forEach(dialog => { dialog.close(); dialog.remove(); });
  });
  window.OfficeImages = Object.freeze({appendImages, validateImage});
})();
