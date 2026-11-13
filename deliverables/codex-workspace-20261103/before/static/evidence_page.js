(() => {
  'use strict';
  window.OfficeEvidenceViewer = {
    async show(api, capture, evidence, parent, current) {
      for(const image of parent.querySelectorAll('img')) window.OfficePdfPreview?.clear(image);
      parent.replaceChildren(); parent.textContent = '正在读取原页…';
      const reply = await api.read_evidence_page(capture, evidence);
      if (!current()) return;
      parent.replaceChildren();
      if (!reply?.ok) { parent.textContent = '原页无法验证或来源已变化，请重新取证。'; return; }
      const data = reply.result, notice = document.createElement('p'), container = document.createElement('div'), image = document.createElement('img');
      notice.textContent = '原始 PDF 第' + data.page + '页；' + data.notice + ' 来源：' + data.source_sha256.slice(0, 12);
      container.className = 'evidence-page-image'; image.alt = '原始 PDF 第' + data.page + '页'; container.append(image);
      if (data.highlight) {
        const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg'), rectangle = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
        svg.setAttribute('viewBox', '0 0 ' + data.width + ' ' + data.height); svg.setAttribute('aria-label', '原文片段边界');
        ['x', 'y', 'width', 'height'].forEach((name, index) => rectangle.setAttribute(name, String(data.highlight[index])));
        rectangle.setAttribute('fill', 'rgba(255,200,0,0.25)'); rectangle.setAttribute('stroke', '#d94d00'); rectangle.setAttribute('stroke-width', '2'); svg.append(rectangle); container.append(svg);
      }
      parent.append(notice, container);
      if(window.OfficePdfPreview){
        const okay = await window.OfficePdfPreview.show(image, data.image, current);
        if(!current())return;
        const retry=document.createElement('button');retry.type='button';retry.className='btn secondary';retry.textContent='重新显示本次原页';
        retry.onclick=async()=>{retry.disabled=true;const okay=await window.OfficePdfPreview.retry(image);if(current()){notice.textContent=okay?'原页已重新显示；来源与授权范围不变。':'原页仍无法显示，请重新取证；未重新下载。';retry.disabled=false;}};
        parent.append(retry);if(!okay)notice.textContent='原页解码未完成，可重新显示；此操作不重新取证或下载。';
      }else image.src=data.image;
    }
  };
})();
