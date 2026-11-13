from types import SimpleNamespace

from tests.test_daily_browser import setup_daily
from tests.test_model_connections import connection,payload,save
from tests.test_professional_browser import _browser
from tests.test_ui_simplify_browser import simplify,pick_model


def test_terminal_feedback_stays_collapsed_while_animation_frame_is_deferred(tmp_path,monkeypatch):
    fixture=_browser(tmp_path,monkeypatch)
    try:
        page,api,original,calls,receipts=next(fixture)
        save(api._connections,payload([connection()]))
        setup_daily(page,api,original,calls,receipts,monkeypatch)
        simplify(page);pick_model(page,'chat','same-model')
        monkeypatch.setattr(api._connections,'build_model',lambda route:SimpleNamespace(invoke=lambda messages:SimpleNamespace(content='摘要应保留研究条件。',response_metadata={'finish_reason':'length'})))
        page.locator('#prompt').fill('说明摘要要求');page.locator('#btn-ask').click()
        page.wait_for_function('document.getElementById("chat-notice").textContent.includes("尚未完整")')
        page.wait_for_function('document.querySelector("[data-turn-notice]")?.hidden')
        page.evaluate('()=>new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve)))')
        hidden=page.evaluate("""async()=>{
          const original=requestAnimationFrame,frames=[],results=[];
          window.requestAnimationFrame=callback=>{frames.push(callback);return frames.length;};
          try {
            for(let i=0;i<3;i++){
              await window.OfficeDaily.refresh();
              results.push([...document.querySelectorAll('[data-turn-notice]')].map(node=>node.hidden));
            }
            return results;
          } finally {
            window.requestAnimationFrame=original;
            for(const callback of frames)original(callback);
          }
        }""")
        assert hidden==[[True],[True],[True]]
        page.evaluate('setStatus("warn","新的操作结果不确定，请先核对，不要重复提交。")')
        page.wait_for_function('!document.getElementById("status-card").hidden')
        assert page.locator('#error-line').is_visible() and '结果不确定' in page.locator('#error-line').inner_text()
        assert calls.count('ask')==1
    finally:fixture.close()
