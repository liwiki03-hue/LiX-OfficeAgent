from copy import deepcopy
import io
from pathlib import Path
import time
import zipfile

import pytest

from office_agent.paged_documents import normalize, parts, projection, sha, W, xml
from tests.test_paged_browser import engine_frame, paged_browser


FIXTURE = Path(__file__).parent / 'fixtures/paged-format-reopen.docx'


def package(values):
    stream = io.BytesIO()
    with zipfile.ZipFile(stream, 'w', zipfile.ZIP_DEFLATED) as archive:
        for name, value in values.items():
            archive.writestr(name, value)
    return stream.getvalue()


@pytest.mark.parametrize('difference', ['none', 'color', 'text', 'explicit_source', 'style', 'tint', 'theme_binding'])
def test_only_unchanged_inherited_color_can_be_restored(difference):
    from lxml import etree as ET
    source_parts = parts(FIXTURE.read_bytes())
    before = xml(source_parts['word/document.xml'])
    original_run = before.find(W+'body').findall(W+'p')[1].find(W+'r')
    if difference == 'explicit_source':
        ET.SubElement(original_run.find(W+'rPr'), W+'color').set(W+'val', '4F81BD')
        source_parts['word/document.xml'] = ET.tostring(before)
    if difference in {'tint', 'theme_binding'}:
        styles = xml(source_parts['word/styles.xml'])
        heading = next(node for node in styles.findall(W+'style') if node.get(W+'styleId') == 'Heading2')
        color = heading.find('./'+W+'rPr/'+W+'color')
        color.set(W+('themeTint' if difference == 'tint' else 'themeColor'), '80' if difference == 'tint' else 'accent2')
        source_parts['word/styles.xml'] = ET.tostring(styles)
    original = package(source_parts)
    candidate_parts = deepcopy(source_parts)
    after = xml(candidate_parts['word/document.xml'])
    run = after.find(W+'body').findall(W+'p')[1].find(W+'r')
    color = run.find('./'+W+'rPr/'+W+'color')
    if color is None:
        color = ET.SubElement(run.find(W+'rPr'), W+'color')
    color.set(W+'val', 'FF0000' if difference == 'color' else '4F81BD')
    if difference == 'text':
        run.find(W+'t').text += '新增正文'
    if difference == 'style':
        styles = xml(candidate_parts['word/styles.xml'])
        heading = next(node for node in styles.findall(W+'style') if node.get(W+'styleId') == 'Heading2')
        heading.find('./'+W+'rPr/'+W+'color').set(W+'val', 'FF0000')
        candidate_parts['word/styles.xml'] = ET.tostring(styles)
    candidate_parts['word/document.xml'] = ET.tostring(after)
    result = normalize(package(candidate_parts), original)
    result_run = xml(parts(result)['word/document.xml']).find(W+'body').findall(W+'p')[1].find(W+'r')
    if difference == 'none':
        assert projection(result) == projection(original)
        assert result_run.find('./'+W+'rPr/'+W+'color') is None
    else:
        assert result_run.find('./'+W+'rPr/'+W+'color') is not None
        if difference == 'explicit_source':
            assert projection(result) == projection(original)
        else:
            assert projection(result) != projection(original)
    assert parts(original) == source_parts


def test_real_engine_reopens_native_formatted_copy_without_freezing_theme_color(paged_browser):
    page, api, source, errors = paged_browser
    original = FIXTURE.read_bytes()
    source.write_bytes(original)
    page.locator('#floating-choose').click()
    page.get_by_role('button', name='同意读取并打开', exact=True).click()
    deadline = time.monotonic() + 60
    while not api._floating.paged.manifest or api._floating.paged.source_version != sha(original):
        if api._floating.paged.reasons:
            raise AssertionError(api._floating.paged.reasons)
        if time.monotonic() > deadline:
            raise AssertionError(page.locator('#paged-feedback').text_content())
        page.wait_for_timeout(100)
    assert api._floating.paged.editable and not api._floating.paged.state()['dirty']
    assert projection(api._floating.paged.data) == projection(original)
    frame = engine_frame(page)
    frame.locator('.superdoc-page').get_by_text('这一段用于核对非目标内容，数字与单位保持。', exact=True).first.click()
    page.keyboard.press('End')
    page.keyboard.insert_text('继续编辑验收。')
    page.get_by_role('button', name='保存副本', exact=True).click()
    page.get_by_role('dialog', name='生成验证副本').get_by_role('button', name='取消', exact=True).click()
    baseline, current = projection(original), projection(api._floating.paged.data)
    assert current['paragraphs'][1] == baseline['paragraphs'][1]
    assert current['paragraphs'][2]['text'] == baseline['paragraphs'][2]['text'] + '继续编辑验收。'
    assert source.read_bytes() == original and not list(source.parent.glob('*.agent-*.docx'))
    assert errors == []


def test_failed_open_releases_layout_for_next_file_read_consent(paged_browser):
    from lxml import etree as ET
    page, api, source, errors = paged_browser
    good = source.read_bytes()
    bad_parts = parts(FIXTURE.read_bytes())
    styles = xml(bad_parts['word/styles.xml'])
    heading = next(node for node in styles.findall(W+'style') if node.get(W+'styleId') == 'Heading2')
    heading.find('./'+W+'rPr/'+W+'color').set(W+'themeTint', '80')
    bad_parts['word/styles.xml'] = ET.tostring(styles)
    source.write_bytes(package(bad_parts))
    page.locator('#floating-choose').click()
    page.get_by_role('button', name='同意读取并打开', exact=True).click()
    page.wait_for_function("document.getElementById('paged-save-state')?.dataset.kind === 'readonly'")
    assert not api._floating.paged.editable
    page.get_by_role('button', name='查看只读原因', exact=True).click()
    reasons = page.get_by_role('dialog', name='只读原因')
    assert reasons.get_by_text(api._floating.paged.reasons[0], exact=True).is_visible()
    reasons.get_by_role('button', name='关闭', exact=True).click()
    page.wait_for_function("document.activeElement.id === 'paged-readonly-reasons'")
    source.write_bytes(good)
    page.locator('#floating-choose').click()
    consent = page.get_by_role('button', name='同意读取并打开', exact=True)
    consent.wait_for()
    page.wait_for_function("!document.getElementById('floating-root').classList.contains('paged-active')")
    assert consent.is_visible() and page.locator('#paged-editor-main').count() == 0
    consent.click()
    deadline = time.monotonic() + 60
    while api._floating.paged.source_version != sha(good) or not api._floating.paged.manifest:
        if time.monotonic() > deadline:
            raise AssertionError(page.locator('#floating-status').text_content())
        page.wait_for_timeout(100)
    assert api._floating.paged.editable and source.read_bytes() == good and errors == []
