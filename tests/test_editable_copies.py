import io
from pathlib import Path
from types import SimpleNamespace
import zipfile

from docx import Document
from docx.oxml import OxmlElement
from docx.shared import Inches
import pytest

from office_agent.editable_copies import body_blocks
from office_agent.instruction_runtime import InstructionRuntime
from office_agent.paged_documents import GuardError, W, eligibility, parts
from office_agent.webui.bridge import WebviewApp
from office_agent.webui.floating_session import FloatingSession
from office_agent.webui.manual_history import checked_output, listing
from office_agent.writes import digest


def make_readonly(source):
    book = Document()
    book.add_heading('开题报告撰写要求', 1)
    paragraph = book.add_paragraph('实验条件为25 ℃，30个样本的延迟为120 ms[1]。')
    paragraph.paragraph_format.tab_stops.add_tab_stop(Inches(1))
    bookmark = OxmlElement('w:bookmarkStart')
    bookmark.set(W + 'id', '7')
    bookmark.set(W + 'name', 'requirements')
    paragraph._p.insert(0, bookmark)
    end = OxmlElement('w:bookmarkEnd')
    end.set(W + 'id', '7')
    paragraph._p.append(end)
    link = OxmlElement('w:hyperlink')
    link.set(W + 'anchor', 'requirements')
    run = OxmlElement('w:r')
    text = OxmlElement('w:t')
    text.text = '链接显示文字'
    run.append(text)
    link.append(run)
    book.add_paragraph()._p.append(link)
    table = book.add_table(rows=2, cols=2)
    table.cell(0, 0).text = '条件'
    table.cell(1, 1).text = '120 ms'
    book.sections[0].header.paragraphs[0].text = '原稿页眉'
    book.save(source)


@pytest.fixture
def readonly(tmp_path):
    source = tmp_path / '只读开题报告.docx'
    make_readonly(source)
    app = WebviewApp()
    api = SimpleNamespace(_app=app, _auth=SimpleNamespace(status=lambda: {'authenticated': True}),
                          _memory_key=lambda: 'self-created-test', _connections=SimpleNamespace(state=SimpleNamespace(mode='agent')))
    floating = FloatingSession(api)
    assert floating.admit(str(source))['ok']
    assert floating.read(floating.read_token, True)['ok']
    service = floating.paged
    assert not service.open(floating.document['id'], floating.document['version'])['editable']
    return service, floating, source


def test_copy_plan_is_explicit_text_only_and_not_a_write(readonly):
    service, floating, source = readonly
    before = source.read_bytes()
    plan = service.plan_editable_copy(service.document_id, 0)['plan']
    assert plan['kind'] == 'paged_editable_copy' and plan['mode'] == 'body_text'
    assert plan['limitations'] and plan['output'].endswith('.docx') and 'candidate' not in plan
    assert not Path(plan['output']).exists() and source.read_bytes() == before
    candidate = service.copy_publisher.plans[plan['id']]['candidate']
    assert eligibility(candidate)['editable']
    assert body_blocks(before) == body_blocks(candidate)
    assert Document(io.BytesIO(candidate)).paragraphs[1].text == '实验条件为25 ℃，30个样本的延迟为120 ms[1]。'
    assert Document(io.BytesIO(candidate)).paragraphs[2].text == '链接显示文字'
    assert not Document(io.BytesIO(candidate)).sections[0].header.paragraphs[0].text
    assert service.publish_editable_copy(plan['id'], service.document_id, 0, False)['cancelled']
    assert not service.copy_publisher.plans and not list(source.parent.glob('*.agent-*.docx'))


def test_copy_confirm_rechecks_requirements_stop_expiry_and_source(readonly):
    service, floating, source = readonly
    plan = service.plan_editable_copy(service.document_id, 0)['plan']
    floating.runtime = InstructionRuntime(floating.document['version'])
    floating.runtime.submit('new-requirement', '保留所有数字')
    with pytest.raises(GuardError, match='要求或会话'):
        service.publish_editable_copy(plan['id'], service.document_id, 0, True)
    plan = service.plan_editable_copy(service.document_id, 0)['plan']
    service.stop()
    with pytest.raises(GuardError, match='过期'):
        service.publish_editable_copy(plan['id'], service.document_id, 0, True)
    plan = service.plan_editable_copy(service.document_id, 0)['plan']
    service.copy_publisher.plans[plan['id']]['deadline'] = 0
    with pytest.raises(GuardError, match='过期'):
        service.publish_editable_copy(plan['id'], service.document_id, 0, True)
    plan = service.plan_editable_copy(service.document_id, 0)['plan']
    source.write_bytes(b'external change')
    with pytest.raises(ValueError, match='外部改变'):
        service.publish_editable_copy(plan['id'], service.document_id, 0, True)
    assert not list(source.parent.glob('*.agent-*.docx'))


@pytest.mark.parametrize('feature', ['protection', 'drawing', 'oMath', 'sym', 'merged', 'vanish', 'fldChar'])
def test_unsafe_copy_is_refused_instead_of_silently_dropping_content(readonly, feature):
    service, floating, source = readonly
    values = parts(source.read_bytes())
    if feature == 'protection':
        values['word/settings.xml'] = values['word/settings.xml'].replace(b'</w:settings>', b'<w:documentProtection w:enforcement="1"/></w:settings>')
    else:
        marker = b'<w:gridSpan w:val="2"/>' if feature == 'merged' else ('<w:' + feature + '/>').encode()
        closing = b'</w:tcPr>' if feature == 'merged' else b'</w:p>'
        values['word/document.xml'] = values['word/document.xml'].replace(closing, marker + closing, 1)
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, 'w', zipfile.ZIP_DEFLATED) as archive:
        for name, data in values.items():
            archive.writestr(name, data)
    with pytest.raises(GuardError):
        body_blocks(buffer.getvalue())


def test_copy_uses_existing_backup_audit_history_and_rollback(readonly, monkeypatch):
    service, floating, source = readonly
    monkeypatch.setattr('office_agent.fidelity.compare_documents', lambda *a, **kw: {'inputs_unchanged': True, 'render': {'status': 'different'}})
    before = source.read_bytes()
    plan = service.plan_editable_copy(service.document_id, 0)['plan']
    publisher = service.copy_publisher
    result = publisher.execute(plan['id'], document_id=service.document_id, draft_revision=0, confirmed=True)
    output = Path(result['output'])
    assert source.read_bytes() == before and Path(result['backup']).read_bytes() == before
    assert result['kind'] == 'paged_editable_copy' and result['status'] == 'verified'
    assert digest(output) == plan['candidate_sha256'] and body_blocks(output.read_bytes()) == body_blocks(before)
    assert listing(floating.app.editor)['records'][0]['transaction_id'] == result['id']
    assert checked_output(floating.app.editor, result['id']) == output
    floating.app.editor.rollback(result['id'], confirmed=True)
    assert not output.exists() and source.read_bytes() == before
