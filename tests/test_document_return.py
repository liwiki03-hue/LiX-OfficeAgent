from copy import deepcopy
import shutil
from threading import RLock
import zipfile

from docx import Document
from lxml import etree as ET
import pytest

from office_agent.document_return import _revision_states
from office_agent.output_dependencies import OutputDependencies
from office_agent.professional_edits import W
from office_agent.webui.document_tasks import DocumentTasks
from office_agent.writes import Editor, digest
from tests.test_professional_edits import fake_render


def test_new_revision_projection_preserves_other_author_and_reports_partial():
    base = ET.fromstring(f'<w:p xmlns:w="{W[1:-1]}"><w:r><w:t>开始</w:t></w:r><w:ins w:id="1" w:author="另一位作者"><w:r><w:t>历史</w:t></w:r></w:ins></w:p>')
    delivered = deepcopy(base)
    for identity, kind, text in [('2', 'del', '旧词'), ('3', 'ins', '新词')]:
        wrapper = ET.SubElement(delivered, W + kind, {W + 'id': identity, W + 'author': 'Office Agent'})
        ET.SubElement(ET.SubElement(wrapper, W + 'r'), W + ('delText' if kind == 'del' else 't')).text = text
    current = deepcopy(delivered)
    current.remove(current[-2])
    result = _revision_states(base, delivered, current)
    assert result['classification'] == 'system_revision_partially_resolved'
    assert result['revision_actions'] == {'2': 'accept', '3': 'keep'}
    current.find(W + 'ins').set(W + 'author', '被改作者')
    assert _revision_states(base, delivered, current) is None


def test_three_version_read_only_return_and_fresh_task(tmp_path, monkeypatch):
    baseline = tmp_path / 'baseline.docx'
    doc = Document()
    for text in ['前置上下文', '旧术语仅限室内20 ms [1]。', '普通用户说明', '结尾上下文']:
        doc.add_paragraph(text)
    doc.save(baseline)
    selected, auth = [baseline], [True]
    editor = Editor(tmp_path, require_confirmation=True, render_validation=True)
    service = DocumentTasks(RLock(), lambda: (auth[0], False, str(selected[0])), lambda: editor, lambda: selected[0], OutputDependencies())
    monkeypatch.setattr('office_agent.advanced_writes.render_gate', fake_render)
    task = service.create('将全文“旧术语”替换为“规范术语”')['result']
    plan = service.plan_group(task['task_id'])['result']
    saved = service.confirm_group(task['task_id'], plan['id'], True)['result']['task']
    delivered = tmp_path / saved['current_name']
    current = tmp_path / 'returned.docx'
    with zipfile.ZipFile(delivered) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    root = ET.fromstring(parts['word/document.xml'])
    paragraphs = root.findall('.//' + W + 'body/' + W + 'p')
    paragraphs[2].find('.//' + W + 't').text = '用户自行补充说明'
    parts['word/document.xml'] = ET.tostring(root)
    with zipfile.ZipFile(current, 'w', zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    hashes = [digest(path) for path in (baseline, delivered, current)]
    selected[0] = current
    payload = {'baseline_name': baseline.name, 'delivered_name': delivered.name, 'receipts': saved['receipts'], 'approved_read': False}
    assert service.compare_return(payload)['result']['status'] == 'cancelled'
    payload['approved_read'] = True
    reply = service.compare_return(payload)
    assert reply['ok'], reply
    result = reply['result']
    assert result['classifications'] == {'unchanged': 2, 'system_suggestion_still_present': 1, 'user_content_change': 1}
    assert result['read_only'] and not result['write_approval_restored']
    assert hashes == [digest(path) for path in (baseline, delivered, current)]
    assert not service.plan_group(task['task_id'])['ok']
    new = service.create('将全文“规范术语”替换为“当前术语”')
    assert new['ok'] and new['result']['source_sha256'] == hashes[-1]
    assert not new['result']['receipts']
    bad = service.compare_return({**payload, 'baseline_name': '../outside.docx'})
    assert bad['reason'] == 'professional_return_filename'
    bad = service.compare_return({**payload, 'receipts': ['0' * 32]})
    assert not bad['ok']
    auth[0] = False
    assert service.compare_return(payload)['reason'] == 'not_authenticated'
