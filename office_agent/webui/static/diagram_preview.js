/* 图表改为对话驱动：不再有类型下拉框或需求输入框。
   本脚本只负责把「最新一次成功渲染」的图片信息交给保存模块（保存/撤销各自单独确认），
   并在登录变化或换文件时回收内存中的预览与导出权限。 */
"use strict";
(() => {
  const authenticated = () => Boolean(window.OfficeAuth?.isAuthenticated());
  let revision = 0;
  function retire() {
    ++revision;
    window.OfficeDiagramExport?.retire?.();
  }
  // 主流程（app.js）在收到 diagram 事件并成功内联渲染后调用这里。
  window.OfficeDiagramPreview = {
    revision: () => revision,
    ready(metadata) {
      if (!authenticated()) return;
      window.OfficeDiagramExport?.ready?.(metadata);
    },
    retire
  };
  ["officeauth:unauthenticated", "officeauth:authenticated", "officediagram:invalidate", "pagehide"]
    .forEach(name => window.addEventListener(name, retire));
  // 用户登出/重置时立即清除私有内容，早于异步鉴权返回。
  ["auth-logout", "auth-reset-confirm"].forEach(id => {
    document.getElementById(id)?.addEventListener("click", retire);
  });
  retire();
})();
