from pathlib import Path
from types import SimpleNamespace
import json

import pytest

from tests.test_chat_context_browser import initialize
from tests.test_daily_flows import finish
from tests.test_professional_browser import _browser, STATIC
from tests.test_chat_results import setting
from tests.test_daily_browser import choose


def extend(page,api,calls):
    names = {'plan_chat_export','confirm_chat_export','open_chat_export_folder','plan_chat_export_rollback','confirm_chat_export_rollback','preview_historical_image','open_historical_image_folder','reuse_historical_image'}
    def invoke(name,args):
        assert name in names
        calls.append(name)
        return getattr(api,name)(*args)
    page.expose_function('resultsHost',invoke)
    page.evaluate('''names => {
      const previous=window.pywebview.api;
      window.pywebview.api=new Proxy({}, {get:(_,name)=>names.includes(name) ? (...args)=>window.resultsHost(name,args) : previous[name]});
    }''',list(names))
    page.add_script_tag(path=str(STATIC/'chat_markdown.js'))


@pytest.mark.parametrize('width,theme',[(1440,'light'),(1280,'dark'),(520,'highcontrast')])
def test_read_continue_export_and_cancel(tmp_path,monkeypatch,width,theme):
    fixture=_browser(tmp_path,monkeypatch)
    try:
        page,api,original,calls,receipts=next(fixture)
        app=initialize(page,api,original,calls,receipts,monkeypatch)
        extend(page,api,calls)
        page.locator('#chat-reply-length').select_option('2048')
        page.locator('#chat-reply-save').click()
        page.wait_for_function('document.getElementById("chat-notice").textContent.includes("回答预算已保存")')
        assert api.preview_chat_context('预算核对')['result']['reply_token_reserve']==2048
        app.window=SimpleNamespace(create_file_dialog=lambda *args,**kwargs:[str(tmp_path)])
        raw='# 标题\n\n- 第一项\n- 第二项\n\n> 引用\n\n```python\nprint("safe")\n```\n\n| 列一 | 列二 |\n| --- | --- |\n| 数值 | 20 ms |\n\n[官网](https://example.org)\n![远程图](https://example.org/private.png)\n<script>window.bad=1</script>\n[危险](javascript:alert(1))'
        requests=[]
        monkeypatch.setattr(api._connections,'build_model',lambda route:SimpleNamespace(invoke=lambda messages:requests.append(messages) or SimpleNamespace(content=raw,response_metadata={'finish_reason':'length'})))
        assert api.ask('保留原问题',{'stream':False})['ok']; finish(app)
        page.evaluate('window.OfficeDaily.refresh()')
        page.wait_for_function('document.querySelector("#chat-transcript h1")')
        page.set_viewport_size({'width':width,'height':800}); page.evaluate('theme=>document.documentElement.dataset.theme=theme',theme)
        assert page.locator('#chat-transcript table td').all_text_contents()==['数值','20 ms']
        assert page.locator('#chat-transcript pre code').text_content()=='print("safe")'
        assert page.locator('#chat-transcript img,#chat-transcript script,#chat-transcript a[href^="javascript:"]').count()==0
        assert page.evaluate('window.bad') is None
        page.get_by_role('button',name='继续回答',exact=True).click()
        page.wait_for_function('document.getElementById("chat-context-content").textContent.includes("关联续写")')
        assert len(requests)==1
        page.locator('#chat-stream').uncheck(); page.locator('#btn-ask').click()
        page.wait_for_function('document.querySelectorAll(".daily-turn").length>=2')
        page.wait_for_function('document.querySelectorAll("#chat-transcript .daily-answer").length>=2')
        assert len(requests)==2 and requests[1][0]['content']=='保留原问题' and requests[1][1]['content']==raw
        page.evaluate('window.approved=false')
        page.get_by_role('button',name='导出此回答',exact=True).first.click()
        page.wait_for_function('document.getElementById("chat-export-result").textContent.includes("已取消")')
        assert not list(tmp_path.glob('.office-chat-export-*'))
        page.evaluate('window.approved=true')
        page.get_by_role('button',name='导出此回答',exact=True).first.click()
        page.wait_for_function('document.getElementById("chat-export-result").textContent.includes("已导出")')
        outputs=list(tmp_path.glob('.office-chat-export-*/conversation.md'))
        assert len(outputs)==1 and raw in outputs[0].read_text(encoding='utf8')
        assert '保留原问题' not in outputs[0].read_text(encoding='utf8')
        assert len(requests)==2
        page.locator('#chat-export-result').scroll_into_view_if_needed()
        assert page.evaluate('document.documentElement.scrollWidth<=innerWidth')
        page.screenshot(path=str(tmp_path/f'chat-results-{width}-{theme}.png'))
    finally:
        fixture.close()


def test_markdown_incomplete_and_stream_append_does_not_steal_scroll(tmp_path,monkeypatch):
    fixture=_browser(tmp_path,monkeypatch)
    try:
        page,api,original,calls,receipts=next(fixture)
        initialize(page,api,original,calls,receipts,monkeypatch); extend(page,api,calls)
        result=page.evaluate(r'''() => {
          const target=document.createElement('div'); document.body.append(target);
          window.OfficeMarkdown.render(target,'```\n<script>alert(1)</script>\n[unclosed](');
          const plain=target.textContent, safe=!target.querySelector('script,img,a'); target.remove(); return {plain,safe};
        }''')
        assert result['safe'] and '<script>' in result['plain']
        page.evaluate('''() => {
          for (let i=0;i<50;i++) { const p=document.createElement('p');p.textContent='历史段落 '+i; document.getElementById('chat-transcript').append(p); }
          const scroller=document.getElementById('workspace-resources');scroller.scrollTop=0;window.savedScroll=scroller.scrollTop;
          window.dispatchEvent(new CustomEvent('officedaily:event',{detail:{type:'answer_draft_start',stream_id:'a'}}));
          for (let i=0;i<20;i++) window.dispatchEvent(new CustomEvent('officedaily:event',{detail:{type:'answer_draft_delta',stream_id:'a',text:'新片段'}}));
        }''')
        assert page.locator('#chat-live-answer').text_content()=='新片段'*20
        assert page.evaluate('document.getElementById("workspace-resources").scrollTop===window.savedScroll')
        assert page.locator('#chat-latest').is_enabled()
    finally:
        fixture.close()


def test_history_preview_reuse_and_missing_file_never_resubmit(tmp_path,monkeypatch):
    import io
    from PIL import Image
    fixture=_browser(tmp_path,monkeypatch)
    try:
        page,api,original,calls,receipts=next(fixture)
        app=initialize(page,api,original,calls,receipts,monkeypatch); extend(page,api,calls)
        setting(api,{'image_reuse_metadata':True})
        choose(page,'image','image-model')
        page.locator('#prompt').fill('画一张蓝色科研插图'); page.locator('#btn-ask').click()
        page.get_by_role('button',name='确认生成（可能产生费用）',exact=True).wait_for()
        _,service=app._image_session
        data=io.BytesIO(); Image.new('RGB',(32,32),'blue').save(data,format='PNG')
        generated=[]; service._provider_factory=lambda:SimpleNamespace(generate=lambda *args:generated.append(args) or data.getvalue())
        page.get_by_role('button',name='确认生成（可能产生费用）',exact=True).click()
        page.get_by_role('dialog',name='确认联网生成图片？').get_by_role('button',name='确认生成（可能产生费用）',exact=True).click()
        page.get_by_role('button',name='保存成新的图片文件…',exact=True).wait_for()
        app.window=SimpleNamespace(create_file_dialog=lambda *args,**kwargs:[str(tmp_path)])
        page.get_by_role('button',name='保存成新的图片文件…',exact=True).click()
        page.get_by_role('button',name='确认保存新文件',exact=True).click()
        page.get_by_role('dialog').get_by_role('button',name='确认保存新文件',exact=True).click()
        page.get_by_role('button',name='撤销这个文件…',exact=True).wait_for()
        page.locator('#daily-open-images').click()
        page.get_by_role('button',name='预览已保存图片',exact=True).click()
        page.wait_for_function('document.querySelector("#daily-image-records img")?.naturalWidth===32')
        page.get_by_role('button',name='复用提示词和参数到草稿',exact=True).click()
        page.wait_for_function('document.getElementById("prompt").value==="画一张蓝色科研插图"')
        assert len(generated)==1
        assert len([name for name in calls if name=='ask'])==1
        page.evaluate('''() => {
          const previous=window.pywebview.api;
          window.pywebview.api=new Proxy({}, {get:(_,name)=>name==='reuse_historical_image' ? (...args)=>previous[name](...args).then(value=>new Promise(resolve=>{window.releaseReuse=()=>resolve(value)})) : previous[name]});
        }''')
        page.get_by_role('button',name='复用提示词和参数到草稿',exact=True).click()
        page.wait_for_function('typeof window.releaseReuse === "function"')
        page.locator('#prompt').fill('新的输入必须保留'); page.evaluate('window.releaseReuse()')
        page.wait_for_function('document.getElementById("daily-image-records").textContent.includes("未覆盖新草稿")')
        assert page.locator('#prompt').input_value()=='新的输入必须保留' and len(generated)==1
        next(tmp_path.glob('office-image-*.png')).unlink()
        page.get_by_role('button',name='预览已保存图片',exact=True).click()
        page.wait_for_function('document.getElementById("daily-image-records").textContent.includes("图片不存在")')
        assert len(generated)==1
        page.locator('#daily-image-records').scroll_into_view_if_needed()
        page.screenshot(path=str(tmp_path/'history-image-reuse.png'))
    except Exception:
        try:
            diagnostic = page.locator('#daily-image-records').evaluate('''node=>({
                images:Array.from(node.querySelectorAll('img')).slice(0,4).map(img=>({
                    complete:img.complete,width:img.naturalWidth,height:img.naturalHeight,
                    loading:img.loading,connected:img.isConnected,top:img.getBoundingClientRect().top,
                    sourcePresent:img.hasAttribute('src')
                })),
                status:Array.from(node.querySelectorAll('[role="status"]')).slice(0,4).map(value=>value.textContent.slice(0,160))
            })''')
            diagnostic['calls'] = {name:calls.count(name) for name in ['ask','preview_historical_image','reuse_historical_image']}
            (tmp_path/'history-preview-diagnostic.json').write_text(json.dumps(diagnostic,ensure_ascii=False,indent=2),encoding='utf8')
        except Exception:
            pass
        raise
    finally:
        fixture.close()


def test_explicit_history_preview_decodes_outside_viewport(tmp_path,monkeypatch):
    import sys
    source_browser = _browser
    def below_viewport(*args, **kwargs):
        fixture = source_browser(*args, **kwargs)
        try:
            context = next(fixture)
            context[0].add_style_tag(content='#daily-image-records article > div { transform:translateY(6000px); }')
            yield context
        finally:
            fixture.close()
    monkeypatch.setattr(sys.modules[__name__], '_browser', below_viewport)
    test_history_preview_reuse_and_missing_file_never_resubmit(tmp_path,monkeypatch)
