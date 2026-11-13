from docx import Document

from office_agent.research_model import ResearchModelError
from office_agent.writes import digest
from test_professional_browser import _browser, STATIC
from test_project_workbench import imported
from test_research_assist import Model


def exercise_completion(tmp_path, monkeypatch, capture=None):
    browser = _browser(tmp_path, monkeypatch)
    try:
        page, api, app, calls, replies = next(browser)
        document = Document()
        for index in range(45):
            document.add_paragraph(f'Owned observation {index}.')
        document.save(app.path)
        original = digest(app.path)
        imported(api)
        anchor = api.read_structure_anchors(40)['result']['anchors'][-1]
        assert api.workbench_action('target', {'locator_id': anchor['locator_id'], 'offset': 40})['ok']
        queued = api.workbench_action('queue', {'operation': 'comment', 'text': 'Recheck the source conditions.', 'evidence_ids': []})
        assert queued['ok'], queued
        allowed = {'workbench_snapshot', 'workbench_action', 'read_structure_anchors'}
        def invoke(method, args):
            assert method in allowed, method
            calls.append((method, args))
            reply = getattr(api, method)(*args)
            replies.append((method, reply))
            return reply
        page.expose_function('completionHost', invoke)
        page.evaluate("window.pywebview.api=new Proxy({}, {get:(_,method)=>(...args)=>window.completionHost(method,args)});void 0")
        page.add_script_tag(path=str(STATIC / 'workbench.js'))
        assert not calls
        page.locator('#wb-load').click()
        page.wait_for_function("document.getElementById('wb-target').options.length===20")
        page.locator('#wb-target').select_option(index=1)
        page.wait_for_function("document.getElementById('wb-original').textContent.includes('observation 1.')")
        page.locator('#wb-queue button').click()
        page.wait_for_function("document.getElementById('wb-original').textContent.includes('observation 44.')")
        assert page.locator('#wb-target').input_value() == anchor['locator_id']
        assert page.locator('#wb-confirm').is_disabled() and not page.locator('#wb-reviewed').is_checked()
        question = 'Do not infer 20 ms in 2024 beyond dry specimens'
        page.locator('#wb-question').fill(question)
        page.locator('#wb-search').click()
        page.wait_for_function("document.querySelectorAll('#wb-gap-summary button').length>=2")
        searches = sum(method == 'workbench_action' and args[0] == 'search' for method, args in calls)
        page.locator('#wb-gap-summary button').first.click()
        assert page.locator('#wb-question').input_value() == question
        assert sum(method == 'workbench_action' and args[0] == 'search' for method, args in calls) == searches == 1
        page.locator('#wb-gap-summary button').filter(has_text='定位登记材料').click()
        page.wait_for_function("document.querySelector('#wb-records input').checked===true")
        assert page.locator('#wb-file').input_value() == ''
        assert not api._workbench.evidence and not api._workbench.captures
        assert '未覆盖页不代表不存在证据' in page.locator('#wb-gap-summary').inner_text()
        if capture:
            capture(page)
        assert digest(app.path) == original
        return {'host_queue_target_beyond_first_page': True, 'gap_focus_keeps_original_question': True,
                'metadata_navigation_without_body_read': True, 'source_unchanged': True,
                'model_calls': 0, 'document_written': False}
    finally:
        browser.close()


def test_gap_material_and_queue_navigation_are_connected_without_approval(tmp_path, monkeypatch):
    exercise_completion(tmp_path, monkeypatch)


def test_failed_research_shows_repair_and_manual_status_does_not_restart_model(tmp_path, monkeypatch):
    browser = _browser(tmp_path, monkeypatch)
    jobs = []
    try:
        page, api, app, calls, replies = next(browser)
        document = Document()
        document.add_paragraph('Latency is 20 ms.')
        document.save(app.path)
        imported(api)
        class Failing(Model):
            def suggest(self, payload, cancelled):
                self.inputs.append(payload)
                raise ResearchModelError('research_model_timeout')
        model = Failing()
        api._research_assist.model_factory = lambda: model
        allowed = {'workbench_snapshot', 'workbench_action', 'read_structure_anchors',
                   'start_workbench_assist', 'research_assist_status'}
        def invoke(method, args):
            assert method in allowed, method
            calls.append((method, args))
            reply = getattr(api, method)(*args)
            if method == 'start_workbench_assist' and reply['ok']:
                jobs.append(api._research_assist.jobs[reply['result']['id']])
            return reply
        page.expose_function('failureHost', invoke)
        page.evaluate("window.pywebview.api=new Proxy({}, {get:(_,method)=>(...args)=>window.failureHost(method,args)});void 0")
        page.add_script_tag(path=str(STATIC / 'workbench.js'))
        page.locator('#wb-load').click()
        page.wait_for_function("!document.getElementById('wb-assist').disabled")
        page.locator('#wb-question').fill('Polish the selected observation.')
        page.locator('#wb-mode').select_option('polish')
        page.locator('#wb-outbound').check()
        page.locator('#wb-assist').click()
        page.wait_for_function("document.getElementById('wb-task-feedback').textContent.includes('超时')")
        assert not page.locator('#wb-outbound').is_checked()
        assert page.locator('#wb-task-queue').is_disabled()
        page.locator('#wb-task-status').click()
        page.wait_for_function("!document.getElementById('wb-task-status').disabled")
        assert len(model.inputs) == 1 and len(jobs) == 1
        assert not api._workbench.live_drafts
        page.locator('#wb-mode').select_option('critique')
        page.locator('#wb-question').fill('Is there sufficient evidence for wet specimens?')
        page.locator('#wb-outbound').check()
        page.locator('#wb-assist').click()
        page.wait_for_function("document.getElementById('wb-task-feedback').textContent.includes('证据不足')")
        assert page.locator('#wb-task-queue').is_disabled()
        assert len(model.inputs) == 1 and len(jobs) == 2
    finally:
        for job in jobs:
            job['cancel'].set()
            job['thread'].join(5)
            assert not job['thread'].is_alive()
        browser.close()
