from docx import Document

from test_professional_browser import _browser, STATIC
from test_project_workbench import imported
from test_project_history import save
from test_research_assist import Model
from office_agent.writes import digest


def exercise_history(tmp_path, monkeypatch, capture=None, late=False):
    browser=_browser(tmp_path,monkeypatch)
    try:
        page,api,app,calls,replies=next(browser)
        doc=Document()
        doc.add_paragraph('The latency is 20 ms.')
        doc.add_paragraph('Unchanged context.')
        doc.save(app.path)
        original=digest(app.path)
        imported(api)
        anchor=api.read_structure_anchors()['result']['anchors'][0]
        assert api.workbench_action('target',{'locator_id':anchor['locator_id'],'offset':0})['ok']
        model=Model()
        api._research_assist.model_factory=lambda:model
        started=api.start_workbench_assist({'question':'Polish the selected wording.','mode':'polish','outbound_approved':True})
        assert started['ok'],started
        job=api._research_assist.jobs[started['result']['id']]
        job['thread'].join(5)
        assert not job['thread'].is_alive() and job['status']=='completed'
        plan=api.plan_workbench_task(job['id'])
        assert plan['ok'],plan
        task=api.confirm_workbench_task(plan['result']['plan_id'],True)
        assert task['ok'],task
        task_receipt=task['result']['save_id']
        first=save(api,'Alpha')
        latest=save(api,'Beta')
        allowed={'workbench_snapshot','workbench_action','read_structure_anchors','restore_project','restore_research_task'}
        def invoke(method,args):
            assert method in allowed,method
            calls.append((method,args))
            reply=getattr(api,method)(*args)
            replies.append((method,reply))
            return reply
        page.expose_function('historyHost',invoke)
        page.evaluate("window.pywebview.api=new Proxy({}, {get:(_,method)=>(...args)=>window.historyHost(method,args).then(reply=>{if(window.delayHistory&&method==='workbench_action'&&args[0]==='project_history')return new Promise(resolve=>window.releaseHistory=()=>{resolve(reply);window.historyDelivered=true});return reply})});void 0")
        page.add_script_tag(path=str(STATIC/'workbench.js'))
        assert calls==[]
        page.locator('#wb-load').click()
        page.wait_for_function("document.getElementById('wb-task-select').options.length===2")
        assert not any(method=='workbench_action' and args[0]=='project_history' for method,args in calls)
        before=len(calls)
        page.locator('#wb-task-select').select_option(task_receipt)
        assert page.locator('#wb-task-receipt').input_value()==task_receipt and len(calls)==before
        page.locator('#wb-task-restore').click()
        page.wait_for_function("document.getElementById('wb-status').textContent.includes('已恢复历史任务供比较')")
        assert len(model.inputs)==1 and not api._workbench.live_drafts
        if late:
            page.evaluate('window.delayHistory=true')
            page.locator('#wb-history-load').click()
            page.wait_for_function("typeof window.releaseHistory==='function'")
            page.evaluate("window.authenticated=false;window.dispatchEvent(new Event('officeauth:unauthenticated'));window.releaseHistory()")
            page.wait_for_function('window.historyDelivered===true')
            assert page.locator('#wb-history-select option').count()==1
            assert page.locator('#wb-task-select option').count()==1
            assert page.locator('#wb-receipt').input_value()==''
            assert digest(app.path)==original
            return {'late_history_ignored':True,'logout_clears_selectors':True}
        page.locator('#wb-name').fill('Unsaved')
        page.locator('#wb-rename').click()
        page.wait_for_function("document.getElementById('wb-project-state').textContent.includes('Unsaved')")
        page.locator('#wb-history-load').click()
        page.wait_for_function("document.getElementById('wb-history-select').options.length===3")
        page.locator('#wb-history-select').select_option(first)
        assert page.locator('#wb-receipt').input_value()==first
        assert api._workbench.project['name']=='Unsaved'
        page.evaluate('window.approved=false')
        before=len(calls)
        page.locator('#wb-restore').click()
        assert len(calls)==before and api._workbench.project['name']=='Unsaved'
        page.evaluate('window.approved=true')
        page.locator('#wb-restore').click()
        page.wait_for_function("document.getElementById('wb-status').textContent.includes('项目已恢复')")
        assert api._workbench.project['name']=='Alpha' and api._workbench.receipt==first
        assert not api._workbench.evidence and not api._workbench.live_drafts
        page.locator('#wb-history-load').click()
        page.wait_for_function("document.getElementById('wb-history-select').options.length===2&&!document.getElementById('wb-history-load').disabled")
        if capture:
            capture(page)
        assert digest(app.path)==original
        return {'project_history_selection':True,'registered_task_selection_and_restore':True,
            'project_receipt':first,'latest_receipt':latest,'task_receipt':task_receipt,
            'no_receipt_manually_copied':True,'no_auto_history_read':True,'cancel_preserves_unsaved_project':True,
            'permissions_restored':False,'document_written':False,'model':'synthetic_stub_only','real_model_calls':0}
    finally:
        browser.close()


def test_registered_history_and_task_selection_restore_without_receipt_copy(tmp_path,monkeypatch):
    exercise_history(tmp_path,monkeypatch)


def test_late_history_response_after_logout_cannot_restore_selectors(tmp_path,monkeypatch):
    exercise_history(tmp_path,monkeypatch,late=True)
