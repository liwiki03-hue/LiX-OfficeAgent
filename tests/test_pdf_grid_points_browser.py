from docx import Document
import pytest

from test_pdf_table import grid_pdf
from test_professional_browser import _browser, STATIC


def exercise_points(tmp_path, monkeypatch, capture=None):
    browser = _browser(tmp_path, monkeypatch)
    try:
        page, api, app, calls, replies = next(browser)
        page.set_viewport_size({'width': 520, 'height': 900})
        document = Document()
        document.add_paragraph('Measured latency is 20 ms.')
        document.save(app.path)
        (tmp_path / 'table.pdf').write_bytes(grid_pdf())
        def invoke(method, args):
            assert method in {'read_structure_anchors', 'read_citation_workspace', 'preview_pdf_numeric', 'read_pdf_numeric'}
            calls.append(method)
            reply = getattr(api, method)(*args)
            replies.append((method, reply))
            return reply
        page.expose_function('pointGridHost', invoke)
        page.evaluate("window.pywebview.api=new Proxy({}, {get:(_,method)=>(...args)=>window.pointGridHost(method,args)});void 0")
        errors = []
        page.on('pageerror', lambda error: errors.append(str(error)))
        page.add_script_tag(path=str(STATIC / 'continuous_research.js'))
        page.locator('#cr-nav-check').click()
        page.locator('#cr-load').click()
        page.wait_for_function("document.getElementById('cr-anchor').options.length===1")
        page.locator('#cr-pdf-table-toggle').click()
        assert page.locator('#cr-pdf-table-point-mode').is_disabled()
        page.locator('#cr-pdf-table-file').fill('table.pdf')
        page.locator('#cr-pdf-table-preview').click()
        page.wait_for_function("!document.getElementById('cr-pdf-table-image').hidden")
        page.wait_for_function("document.getElementById('cr-pdf-table-image').naturalWidth>0")
        image = page.locator('#cr-pdf-table-image')
        def point(x, y):
            box = image.bounding_box()
            image.click(position={'x': box['width'] * x / 100, 'y': box['height'] * y / 100})
        point(10, 10)
        assert page.locator('#cr-pdf-table-rows').input_value() == ''
        before = len(calls)
        for mode, points in [('rows', [(10, 62), (10, 16), (10, 35)]), ('columns', [(98, 10), (2, 10), (42, 10)])]:
            page.locator('#cr-pdf-table-point-mode').select_option(mode)
            for x, y in points:
                point(x, y)
        row_text = page.locator('#cr-pdf-table-rows').input_value()
        column_text = page.locator('#cr-pdf-table-columns').input_value()
        assert [float(v) for v in row_text.split(',')] == pytest.approx([16, 35, 62], abs=.5)
        assert [float(v) for v in column_text.split(',')] == pytest.approx([2, 42, 98], abs=.5)
        assert len(calls) == before
        point(42, 10)
        assert '已存在' in page.locator('#cr-status').inner_text()
        page.locator('#cr-pdf-table-read').click()
        page.wait_for_function("!document.getElementById('cr-pdf-table-ack').disabled")
        assert page.locator('#cr-pdf-table-cells').inner_text().find('20') >= 0
        page.locator('#cr-pdf-table-point-mode').select_option('rows')
        point(10, 80)
        assert page.locator('#cr-pdf-table-ack').is_disabled()
        assert page.locator('#cr-pdf-table-cells').inner_text() == ''
        page.locator('#cr-pdf-table-point-undo').click()
        assert page.locator('#cr-pdf-table-rows').input_value() == row_text
        assert page.locator('#cr-pdf-table-ack').is_disabled()
        assert page.locator('#cr-observation option').count() == 0
        page.locator('#cr-pdf-table-units').fill('2:1')
        assert page.locator('#cr-pdf-table-point-undo').is_disabled()
        page.locator('#cr-pdf-table-point-mode').select_option('columns')
        point(70, 10)
        assert page.locator('#cr-pdf-table-units').input_value() == ''
        page.locator('#cr-pdf-table-point-undo').click()
        assert page.locator('#cr-pdf-table-units').input_value() == '2:1'
        assert page.locator('#cr-pdf-table-ack').is_disabled()
        page.locator('#cr-pdf-table-point-clear').click()
        assert page.locator('#cr-pdf-table-columns').input_value() == ''
        page.locator('#cr-pdf-table-point-undo').click()
        assert page.locator('#cr-pdf-table-columns').input_value() == column_text
        page.locator('#cr-pdf-table-columns').fill('2,42,98')
        assert page.locator('#cr-pdf-table-units').input_value() == ''
        page.evaluate("""window.pywebview.api=new Proxy({}, {get:(_,method)=>(...args)=>window.pointGridHost(method,args).then(reply=>{
            if(method!=='read_pdf_numeric')return reply;
            return new Promise(resolve=>window.releasePointGrid=()=>resolve(reply));
        })});void 0""")
        page.locator('#cr-pdf-table-read').click()
        page.wait_for_function("typeof window.releasePointGrid==='function'")
        held = page.locator('#cr-pdf-table-columns').input_value()
        page.evaluate("const n=document.getElementById('cr-pdf-table-image'),b=n.getBoundingClientRect();n.dispatchEvent(new MouseEvent('click',{clientX:b.left+b.width*.8,clientY:b.top+b.height*.1}));")
        assert page.locator('#cr-pdf-table-columns').input_value() == held
        page.evaluate('window.releasePointGrid()')
        page.wait_for_function("!document.getElementById('cr-pdf-table-ack').disabled")
        if capture:
            capture(page)
        page.locator('#cr-pdf-table-rows').fill('62,16')
        page.locator('#cr-pdf-table-point-mode').select_option('rows')
        point(10, 35)
        assert page.locator('#cr-pdf-table-rows').input_value() == '62,16'
        assert '已有边界无效' in page.locator('#cr-status').inner_text()
        page.locator('#cr-pdf-table-rows').fill('')
        page.locator('#cr-pdf-table-columns').fill('')
        for value in range(5, 21):
            point(10, value)
        point(10, 21)
        assert '16次' in page.locator('#cr-status').inner_text()
        assert len(page.locator('#cr-pdf-table-rows').input_value().split(',')) == 16
        page.locator('#cr-pdf-table-point-undo').click()
        point(10, 21)
        assert float(page.locator('#cr-pdf-table-rows').input_value().split(',')[-1]) == pytest.approx(21, abs=.5)
        page.locator('#cr-pdf-table-page').fill('2')
        assert page.locator('#cr-pdf-table-point-undo').is_disabled()
        assert page.locator('#cr-pdf-table-point-mode').input_value() == 'inspect'
        assert image.is_hidden()
        page.evaluate("window.authenticated=false;window.dispatchEvent(new Event('officeauth:unauthenticated'))")
        assert page.locator('#cr-pdf-table-rows').input_value() == ''
        assert page.locator('#cr-pdf-table-point-undo').is_disabled()
        assert not errors
        return {'point_selection_sorted': True, 'duplicate_noop': True, 'no_implicit_host_calls': True,
                'undo_does_not_restore_layout_approval': True, 'columns_retire_units': True,
                'busy_guard': True, 'history_limit': 16, 'invalid_edges_preserved_for_correction': True,
                'source_and_logout_retire_history': True, 'errors': errors}
    finally:
        browser.close()


def test_pdf_grid_point_selection_and_retirement(tmp_path, monkeypatch):
    exercise_points(tmp_path, monkeypatch)
