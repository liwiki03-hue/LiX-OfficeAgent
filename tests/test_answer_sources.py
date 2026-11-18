import base64
from copy import deepcopy
import io
import threading
import zipfile
from types import SimpleNamespace

from docx import Document
import pytest

from office_agent.webui.bridge import WebviewApp
from office_agent.webui.floating_session import FloatingSession
from tests.test_document_conversation import Answers, FixedAnswer, attach, finish, send
from tests.test_editable_copies import readonly


TEXT = '原文引用“正文第99段”。实验温度25 ℃，30个样本的平均延迟120 ms[1]。'


@pytest.fixture
def answer_document(tmp_path):
    source = tmp_path / '来源位置样例.docx'
    book = Document()
    for text in ('来源定位样例', '非目标段落314。', TEXT, TEXT):
        book.add_paragraph(text)
    book.save(source)
    app = WebviewApp()
    api = SimpleNamespace(_app=app, _auth=SimpleNamespace(status=lambda: {'authenticated':True}),
                          _memory_key=lambda: 'self-created-source-test', _connections=SimpleNamespace(state=SimpleNamespace(mode='agent')))
    floating = FloatingSession(api)
    assert floating.admit(str(source))['ok']
    assert floating.read(floating.read_token, True)['ok']
    service = floating.paged
    assert service.open(floating.document['id'], floating.document['version'])['editable']
    manifest = [{'id':'node-' + str(index), 'text':paragraph.text} for index, paragraph in enumerate(book.paragraphs)]
    service.sync(service.document_id, 0, base64.b64encode(source.read_bytes()).decode(), manifest, True)
    reference = service.reference(service.document_id, 0, {'node_id':'node-3', 'start':0, 'end':len(TEXT)})['reference']
    return floating, service, source, reference


def test_host_uses_actual_node_for_repeated_selection_without_new_request(answer_document):
    floating, service, source, reference = answer_document
    connections = FixedAnswer({'content':'平均延迟为120 ms。'})
    attach(floating, connections)
    original, draft = source.read_bytes(), service.data
    send(floating, 'actual-node', '平均延迟是多少？请附简短来源。', reference=reference['id'])
    finish(floating)
    row = floating.conversation.turns['actual-node']
    assert row['status'] == 'complete' and row['answer'] == '平均延迟为120 ms。\n（已读范围：正文第4段）'
    assert '正文第4段' in connections.calls[0] and len(connections.calls) == 1
    assert floating.conversation.answer_provenance['actual-node']['unique']
    assert row['references'] == [{'scope':'仅当前引用', 'draft_revision':0}]
    assert source.read_bytes() == original and service.data == draft
    assert not service.proposals and not floating.outputs and not floating.access.state()['id']


def test_wide_duplicate_material_lists_read_positions_instead_of_guessing_first(answer_document):
    floating, _, source, _ = answer_document
    connections = FixedAnswer({'content':'平均延迟为120 ms。'})
    attach(floating, connections)
    original = source.read_bytes()
    send(floating, 'wide-material', '平均延迟是多少，单位是什么？')
    finish(floating)
    row = floating.conversation.turns['wide-material']
    assert row['status'] == 'complete'
    assert row['answer'].endswith('（已读范围：正文第3段、正文第4段）')
    assert not floating.conversation.answer_provenance['wide-material']['unique']
    assert '正文第1段' not in row['answer']
    assert len(connections.calls) == 1 and source.read_bytes() == original and not floating.outputs


@pytest.mark.parametrize('mutation', ['caller_context', 'caller_location', 'host_locator'])
def test_unbound_reference_fields_fail_before_model(answer_document, mutation):
    floating, service, source, reference = answer_document
    changed = deepcopy(reference)
    if mutation == 'caller_context':
        changed['context'] = '伪造平均延迟999 ms。'
    elif mutation == 'caller_location':
        changed['location'] = '正文第1段'
    else:
        service.references[reference['id']]['draft_locator_id'] = 'draftloc:forged'
        changed = deepcopy(service.references[reference['id']])
    original = source.read_bytes()
    with pytest.raises(ValueError, match='引用|定位'):
        floating.conversation._material('平均延迟是多少？', changed, with_sources=True)
    assert source.read_bytes() == original and not floating.outputs


def test_stale_reference_cannot_fall_back_to_current_text(answer_document):
    floating, service, source, reference = answer_document
    original = source.read_bytes()
    service.changed(service.document_id, 1)
    with pytest.raises(ValueError, match='最新草稿尚未同步'):
        floating.conversation._material('平均延迟是多少？', reference, with_sources=True)
    assert source.read_bytes() == original and not floating.outputs


@pytest.mark.parametrize('metadata', ['来源：正文第99段[1]', '来源：**正文第99段**'])
def test_explicit_invalid_source_metadata_is_rejected_once(answer_document, metadata):
    floating, _, source, reference = answer_document
    connections = FixedAnswer({'content':'平均延迟为120 ms。\n' + metadata})
    attach(floating, connections)
    original = source.read_bytes()
    send(floating, 'invalid-metadata', '平均延迟是多少？请附来源。', reference=reference['id'])
    finish(floating)
    row = floating.conversation.turns['invalid-metadata']
    assert row['status'] == 'failed' and not row['answer']
    assert '来源定位' in row['text'] and len(connections.calls) == 1
    assert 'invalid-metadata' not in floating.conversation.answer_provenance
    assert source.read_bytes() == original and not floating.outputs


def test_paragraph_mentions_in_normal_explanation_are_not_source_metadata(answer_document):
    floating, _, _, reference = answer_document
    response = '“正文第99段”是原文中的交叉引用；平均延迟为120 ms。'
    connections = FixedAnswer({'content':response})
    attach(floating, connections)
    send(floating, 'quoted-position', '原文引用的“正文第99段”是什么意思？', reference=reference['id'])
    finish(floating)
    row = floating.conversation.turns['quoted-position']
    assert row['status'] == 'complete' and response in row['answer']
    assert row['answer'].endswith('（已读范围：正文第4段）') and len(connections.calls) == 1


def test_long_requested_answer_is_preserved_without_clipping(answer_document):
    floating, _, _, reference = answer_document
    response = '解释：' + '保持原文条件与单位。' * 160
    connections = FixedAnswer({'content':response})
    attach(floating, connections)
    send(floating, 'long-explanation', '请详细解释选区内容并附来源。', reference=reference['id'])
    finish(floating)
    row = floating.conversation.turns['long-explanation']
    assert row['status'] == 'complete' and row['answer'].startswith(response)
    assert len(row['answer']) > len(response) and len(connections.calls) == 1


def test_same_revision_snapshot_change_discards_old_source_answer(answer_document):
    floating, service, source, reference = answer_document
    entered, release = threading.Event(), threading.Event()
    connections = Answers(entered, release)
    attach(floating, connections)
    original = source.read_bytes()
    try:
        send(floating, 'snapshot-change', '平均延迟是多少？', reference=reference['id'])
        assert entered.wait(3)
        book = Document(io.BytesIO(service.data))
        book.paragraphs[3].text = '后来的人工内容125 ms。'
        data = io.BytesIO()
        book.save(data)
        with service.lock:
            service.data = data.getvalue()
    finally:
        release.set()
        finish(floating)
    row = floating.conversation.turns['snapshot-change']
    assert row['status'] == 'failed' and not row['answer']
    assert '快照已改变' in row['text'] and len(connections.calls) == 1
    assert source.read_bytes() == original and not floating.outputs


def test_zip_container_only_change_keeps_same_revision_answer(answer_document):
    floating, service, source, reference = answer_document
    entered, release = threading.Event(), threading.Event()
    connections = Answers(entered, release)
    attach(floating, connections)
    original = source.read_bytes()
    before = service._snapshot_identity(service.data)
    raw_before = service.data
    try:
        send(floating, 'container-change', '平均延迟是多少？', reference=reference['id'])
        assert entered.wait(3)
        data = io.BytesIO(service.data)
        with zipfile.ZipFile(data, 'a') as archive:
            archive.comment = b'self-created container-only marker'
        with service.lock:
            service.data = data.getvalue()
        assert service.data != raw_before and service._snapshot_identity(service.data) == before
    finally:
        release.set()
        finish(floating)
    row = floating.conversation.turns['container-change']
    assert row['status'] == 'complete' and row['answer'].endswith('（已读范围：正文第4段）')
    assert len(connections.calls) == 1 and source.read_bytes() == original and not floating.outputs


def test_long_decimal_source_fragments_are_not_rounded_into_same_location():
    from office_agent.webui.document_conversation import DocumentConversation
    evidence = {'sources': [
        {'label':'正文第2段', 'text':'123456789012345678901234567890 ms', 'aliases':[]},
        {'label':'正文第3段', 'text':'123456789012345678901234567891 ms', 'aliases':[]},
    ]}
    answer, provenance = DocumentConversation._answer_scope('123456789012345678901234567891 ms。', evidence)
    assert answer.endswith('（已读范围：正文第3段）')
    assert provenance['labels'] == ['正文第3段'] and provenance['unique']


def test_exact_host_table_label_does_not_require_new_position_grammar(readonly):
    _, floating, source = readonly
    response = '条件。\n已读范围：表格1行1'
    connections = FixedAnswer({'content': response})
    attach(floating, connections)
    original = source.read_bytes()
    send(floating, 'legacy-table-label', '表格第一行写了什么？请附来源。')
    finish(floating)
    row = floating.conversation.turns['legacy-table-label']
    assert row['status'] == 'complete' and row['answer'] == response
    assert len(connections.calls) == 1 and source.read_bytes() == original and not floating.outputs
