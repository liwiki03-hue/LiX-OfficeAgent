from docx import Document
import pytest

from office_agent.writes import digest
from test_professional_browser import _browser, STATIC
from test_project_history import save
from test_project_workbench import imported
from test_research_assist import Model


def exercise_task_scope(tmp_path, monkeypatch, change='new', capture=None):
    browser = _browser(tmp_path, monkeypatch)
    jobs = []
    try:
        page, api, app, calls, replies = next(browser)
        document = Document()
        document.add_paragraph('The latency is 20 ms.')
        document.add_paragraph('The second observation remains unchanged.')
        document.save(app.path)
        original = digest(app.path)
        imported(api)
        anchor = api.read_structure_anchors()['result']['anchors'][0]
        assert api.workbench_action('target', {'locator_id': anchor['locator_id'], 'offset': 0})['ok']
        receipt = save(api, 'Task context fixture')
        model = Model()
        api._research_assist.model_factory = lambda: model
        allowed = {'workbench_snapshot', 'workbench_action', 'read_structure_anchors', 'start_workbench_assist',
                   'research_assist_status', 'plan_workbench_task', 'confirm_workbench_task',
                   'cancel_research_assist', 'restore_project', 'restore_research_task'}
        def invoke(method, args):
            assert method in allowed, method
            calls.append((method, args))
            reply = getattr(api, method)(*args)
            replies.append((method, reply))
            if method == 'start_workbench_assist' and reply['ok']:
                jobs.append(api._research_assist.jobs[reply['result']['id']])
            return reply
        page.expose_function('taskScopeHost', invoke)
        page.evaluate("window.pywebview.api=new Proxy({}, {get:(_,method)=>(...args)=>window.taskScopeHost(method,args).then(reply=>{if(window.delayTaskStatus&&method==='research_assist_status'&&!window.releaseTaskStatus)return new Promise(resolve=>window.releaseTaskStatus=()=>{resolve(reply);window.taskStatusReleased=true});return reply})});void 0")
        page.add_script_tag(path=str(STATIC / 'workbench.js'))
        assert calls == []
        page.locator('#wb-load').click()
        page.wait_for_function("document.getElementById('wb-target').options.length===2&&!document.getElementById('wb-assist').disabled")
        page.locator('#wb-question').fill('Polish the first observation.')
        page.locator('#wb-mode').select_option('polish')
        page.locator('#wb-outbound').check()
        if change in {'late', 'logout'}:
            page.evaluate('window.delayTaskStatus=true')
        page.locator('#wb-assist').click()
        if change in {'late', 'logout'}:
            page.wait_for_function("typeof window.releaseTaskStatus==='function'")
            assert all(page.locator('#wb-' + name).is_disabled() for name in ['new', 'restore', 'target', 'task-restore'])
            if change == 'logout':
                page.evaluate("window.authenticated=false;window.dispatchEvent(new Event('officeauth:unauthenticated'));window.releaseTaskStatus()")
                page.wait_for_function('window.taskStatusReleased===true')
                assert page.locator('#wb-task').inner_text() == ''
                assert page.locator('#wb-task-confirm').is_disabled() and page.locator('#wb-task-queue').is_disabled()
                assert not page.locator('#wb-outbound').is_checked()
                assert digest(app.path) == original and len(model.inputs) == 1
                return {'logout_retires_late_task_status': True, 'real_model_calls': 0,
                        'synthetic_calls': 1, 'source_unchanged': True}
            old_id = jobs[0]['id']
            page.locator('#wb-cancel-task').click()
            page.wait_for_function("document.getElementById('wb-task').textContent.includes('user_cancelled')")
            page.locator('#wb-question').fill('Polish the second bounded request.')
            page.locator('#wb-assist').click()
            page.wait_for_function("!document.getElementById('wb-task-save').disabled")
            new_id = jobs[-1]['id']
            assert new_id != old_id
            page.evaluate('window.releaseTaskStatus()')
            page.wait_for_function('window.taskStatusReleased===true')
            assert new_id in page.locator('#wb-task').inner_text()
            assert old_id not in page.locator('#wb-task').inner_text()
            assert len(model.inputs) == 2
            assert digest(app.path) == original
            return {'old_poll_cannot_replace_new_task': True, 'running_context_switch_disabled': True,
                    'real_model_calls': 0, 'synthetic_calls': len(model.inputs), 'source_unchanged': True}
        page.wait_for_function("!document.getElementById('wb-task-save').disabled")
        page.locator('#wb-task-save').click()
        page.wait_for_function("!document.getElementById('wb-task-confirm').disabled")
        identifier = next(reply['result']['plan_id'] for method, reply in reversed(replies) if method == 'plan_workbench_task')
        if change == 'new':
            old_project = api._workbench.project['id']
            page.locator('#wb-new').click()
            page.wait_for_function("document.getElementById('wb-task').textContent===''")
            assert api._workbench.project['id'] != old_project
        elif change == 'target':
            page.locator('#wb-target').select_option(index=1)
            page.wait_for_function("document.getElementById('wb-original').textContent.includes('second observation')")
        elif change == 'restore':
            assert page.locator('#wb-receipt').input_value() == receipt
            page.locator('#wb-restore').click()
            page.wait_for_function("document.getElementById('wb-status').textContent.includes('项目已恢复')")
        else:
            page.locator('#wb-task-confirm').click()
            page.wait_for_function("document.getElementById('wb-status').textContent.includes('任务已保存并登记')")
            page.locator('#wb-task-save').click()
            page.wait_for_function("!document.getElementById('wb-task-confirm').disabled")
            page.locator('#wb-task-restore').click()
            page.wait_for_function("document.getElementById('wb-status').textContent.includes('已恢复历史任务供比较')")
        assert page.locator('#wb-task-confirm').is_disabled()
        assert page.locator('#wb-task-queue').is_disabled()
        assert page.locator('#wb-task-save').is_disabled()
        assert not page.locator('#wb-outbound').is_checked()
        if change != 'history':
            assert page.locator('#wb-task').inner_text() == ''
            assert identifier not in api._research_assist._task_plans
            assert not api.confirm_workbench_task(identifier, True)['ok']
        if capture:
            capture(page)
        assert digest(app.path) == original and not api._workbench.live_drafts
        assert len(model.inputs) == 1
        return {'context': change, 'task_preview_retired': True, 'old_task_not_adoptable': True,
                'outbound_reapproval_required': True, 'source_unchanged': True, 'document_written': False,
                'real_model_calls': 0, 'synthetic_calls': len(model.inputs),
                'project_receipt': receipt, 'old_task_id': jobs[0]['id']}
    finally:
        for job in jobs:
            job['cancel'].set()
            job['thread'].join(5)
            assert not job['thread'].is_alive()
        browser.close()


@pytest.mark.parametrize('change', ['new', 'target', 'restore', 'history'])
def test_context_switch_retires_task_and_save_preview(tmp_path, monkeypatch, change):
    exercise_task_scope(tmp_path, monkeypatch, change)


@pytest.mark.parametrize('change', ['late', 'logout'])
def test_late_task_poll_cannot_replace_current_view(tmp_path, monkeypatch, change):
    exercise_task_scope(tmp_path, monkeypatch, change)
