import pytest

from tests.test_professional_browser import STATIC,_browser


@pytest.mark.parametrize('action',['Escape','取消','确认生成（可能产生费用）'])
def test_image_confirmation_settles_without_waiting_for_close_event(tmp_path,monkeypatch,action):
    fixture=_browser(tmp_path,monkeypatch)
    try:
        page,_,_,calls,_=next(fixture)
        page.add_script_tag(path=str(STATIC/'imagegen.js'))
        page.evaluate('''() => {
          const card=document.createElement('article');document.body.append(card);window.imageCalls=[];
          OfficeImageGen.handleEvent({type:'image_plan',plan:{plan_id:'test-plan',effective_prompt:'同步取消验收',n:1}},
            async(...args)=>{imageCalls.push(args);return {ok:true};},node=>card.append(node));
        }''')
        page.get_by_role('button',name='确认生成（可能产生费用）',exact=True).click()
        dialog=page.get_by_role('dialog',name='确认联网生成图片？')
        dialog.evaluate('''node=>{
          node.addEventListener('close',event=>event.stopImmediatePropagation(),true);
          window.oldConfirm=node.querySelector('button:last-child');
        }''')
        if action=='Escape':page.keyboard.press('Escape')
        else:dialog.get_by_role('button',name=action,exact=True).click()
        assert page.locator('.imagegen-confirm').count()==0
        if action.startswith('确认'):
            page.wait_for_function('imageCalls.length===1')
            page.evaluate('window.oldConfirm.click()')
            page.wait_for_timeout(40)
            assert page.evaluate('imageCalls')==[['confirm_image_generation','test-plan',True]]
        else:
            assert page.evaluate('imageCalls')==[]
        assert not calls
        page.screenshot(path=str(tmp_path/'dialog-settled.png'))
    finally:
        fixture.close()
