from copy import deepcopy
from threading import RLock
import zipfile

from docx import Document
from lxml import etree as ET
import pytest

from office_agent.document_return import _normalized, _project
from office_agent.output_dependencies import OutputDependencies
from office_agent.professional_edits import W, _canonical
from office_agent.webui.document_tasks import DocumentTasks
from office_agent.writes import Editor, digest
from tests.test_professional_edits import fake_render


def parts(path):
    with zipfile.ZipFile(path) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def rewrite(path, action):
    contents = parts(path)
    root = ET.fromstring(contents['word/document.xml'])
    action(root.findall('.//' + W + 'body/' + W + 'p'))
    contents['word/document.xml'] = ET.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as archive:
        for name, data in contents.items():
            archive.writestr(name, data)


def prepared(tmp_path, monkeypatch, *, real=False):
    baseline = tmp_path / 'baseline.docx'
    document = Document()
    for text in ['前置上下文', '旧术语仅限室内20 ms [1]。', '普通用户说明', '重复重复重复', '后续核对区域。', '尾部上下文']:
        document.add_paragraph(text)
    table = document.add_table(rows=1, cols=1)
    table.cell(0, 0).text = '表格候选保留原入口'
    document.save(baseline)
    def history(paragraphs):
        wrapper = ET.SubElement(paragraphs[4], W + 'ins', {W + 'id': '42', W + 'author': '另一位作者', W + 'date': '2026-01-01T00:00:00Z'})
        ET.SubElement(ET.SubElement(wrapper, W + 'r'), W + 't').text = '其他作者保留'
    rewrite(baseline, history)
    selected, authenticated = [baseline], [True]
    editor = Editor(tmp_path, require_confirmation=True, render_validation=True)
    service = DocumentTasks(RLock(), lambda: (authenticated[0], False, str(selected[0])), lambda: editor, lambda: selected[0], OutputDependencies())
    if not real:
        monkeypatch.setattr('office_agent.advanced_writes.render_gate', fake_render)
    task = service.create('将全文“旧术语”替换为“规范术语”')['result']
    plan = service.plan_group(task['task_id'])['result']
    task = service.confirm_group(task['task_id'], plan['id'], True)['result']['task']
    delivered = tmp_path / task['current_name']
    current = tmp_path / 'returned.docx'
    current.write_bytes(delivered.read_bytes())
    rewrite(current, lambda paragraphs: setattr(paragraphs[2].find('.//' + W + 't'), 'text', '用户自行补充说明'))
    selected[0] = current
    service.invalidate()
    payload = {'baseline_name': baseline.name, 'delivered_name': delivered.name, 'receipts': task['receipts'], 'approved_read': True}
    compared = service.compare_return(payload)
    assert compared['ok'], compared
    return baseline, delivered, current, editor, service, selected, authenticated, payload, compared['result']


def request(result, text='规范术语', replacement='本轮术语', **kwargs):
    row = next(row for row in result['items'] if text in row['current'])
    return {'comparison_id': result['comparison_id'], 'row_id': row['continuation_id'], 'find': text,
            'replacement': replacement, 'reason': '仅调整术语，保留20 ms及引文和其他作者', 'acknowledge_current': True, **kwargs}


@pytest.mark.parametrize('mode', ['clean', 'tracked'])
def test_return_new_task_cancel_copy_projection_and_undo(tmp_path, monkeypatch, mode):
    baseline, delivered, current, editor, service, _, _, payload, result = prepared(tmp_path, monkeypatch)
    originals = [digest(path) for path in (baseline, delivered, current)]
    count = len(list(editor.audit.iterdir()))
    draft = request(result, text='后续核对', replacement='继续核对', mode=mode)
    reply = service.continue_return(draft)
    assert reply['ok'], reply
    task = reply['result']
    assert task['source_sha256'] == originals[-1] and task['receipts'] == [] and task['wrote_nothing']
    assert not task['write_approval_restored'] and not task['rollback_authority_restored']
    assert len(list(editor.audit.iterdir())) == count
    assert not service.continue_return(draft)['ok']
    plan = service.plan_group(task['task_id'])['result']
    assert service.confirm_group(task['task_id'], plan['id'], False)['result']['status'] == 'cancelled'
    assert len(list(editor.audit.iterdir())) == count
    plan = service.plan_group(task['task_id'])['result']
    saved = service.confirm_group(task['task_id'], plan['id'], True)
    assert saved['ok'], saved
    output = tmp_path / saved['result']['task']['current_name']
    before, after = parts(current), parts(output)
    assert all(after[name] == data for name, data in before.items() if name != 'word/document.xml')
    old_root, new_root = [ET.fromstring(data['word/document.xml']) for data in (before, after)]
    old_history = next(node for node in old_root.iter(W + 'ins') if node.get(W + 'id') == '42')
    assert _canonical(old_history) == _canonical(next(node for node in new_root.iter(W + 'ins') if node.get(W + 'id') == '42'))
    assert '用户自行补充说明' in ''.join(new_root.itertext())
    if mode == 'tracked':
        old_paragraph = old_root.findall('.//' + W + 'body/' + W + 'p')[4]
        new_paragraph = new_root.findall('.//' + W + 'body/' + W + 'p')[4]
        ids = [node.get(W + 'id') for node in new_paragraph if node.tag in {W + 'ins', W + 'del'} and node.get(W + 'id') != '42']
        assert len(ids) == 2
        assert _normalized(_project(new_paragraph, ids, ['reject'] * len(ids))) == _normalized(old_paragraph)
        approved = deepcopy(old_paragraph)
        approved.find(W + 'r/' + W + 't').text = '继续核对区域。'
        assert _normalized(_project(new_paragraph, ids, ['accept'] * len(ids))) == _normalized(approved)
    else:
        assert '继续核对区域。' in ''.join(new_root.itertext())
    assert originals == [digest(path) for path in (baseline, delivered, current)]
    rolled = service.undo_last(task['task_id'], True)
    assert rolled['ok'] and rolled['result']['status'] == 'rolled_back' and not output.exists()
    assert originals == [digest(path) for path in (baseline, delivered, current)]


@pytest.mark.parametrize('which', [0, 1, 2])
def test_any_of_three_versions_changed_rejects_continuation(tmp_path, monkeypatch, which):
    baseline, delivered, current, editor, service, _, _, _, result = prepared(tmp_path, monkeypatch)
    count = len(list(editor.audit.iterdir()))
    path = [baseline, delivered, current][which]
    path.write_bytes(path.read_bytes() + b' ')
    assert service.continue_return(request(result))['reason'] == 'professional_source_changed'
    assert len(list(editor.audit.iterdir())) == count


def test_ambiguous_conflict_sensitive_and_protected_targets_are_not_replayed(tmp_path, monkeypatch):
    baseline, delivered, current, editor, service, _, _, payload, result = prepared(tmp_path, monkeypatch)
    row = next(row for row in result['items'] if row['classification'] == 'user_content_change')
    assert row['continuation_id'] and row['continuation_requires_current_review'] and row['requires_review']
    assert next(row for row in result['items'] if '/table/' in (row['current_location'] or ''))['continuation_id'] is None
    assert service.continue_return(request(result, text='重复', replacement='唯一'))['reason'] == 'professional_return_unique_text'
    assert service.continue_return(request(result, text='重复重复', replacement='唯一'))['reason'] == 'professional_return_unique_text'
    assert service.continue_return(request(result, text='其他作者保留', replacement='替换作者内容'))['reason'] == 'professional_protected_boundary'
    assert not service.continue_return(request(result, text='20', replacement='21'))['ok']
    assert not service.continue_return({**request(result), 'path': str(current)})['ok']
    assert not service.continue_return({**request(result), 'acknowledge_current': False})['ok']
    sensitive = service.continue_return({**request(result), 'reason': 'password=synthetic-fixture'})
    assert sensitive == {'ok': False, 'reason': 'professional_task_request'}
    def conflicting(paragraphs):
        paragraphs[1].find('.//' + W + 't').text = '用户冲突内容20 ms [1]。'
        paragraphs[2].find('.//' + W + 't').text = '普通用户说明'
    rewrite(current, conflicting)
    compared = service.compare_return(payload)['result']
    conflict = next(row for row in compared['items'] if row['classification'] == 'same_position_conflict')
    assert conflict['continuation_id'] is None


@pytest.mark.parametrize('action', ['expire', 'logout', 'invalidate', 'recompare', 'cancel'])
def test_lifecycle_invalidates_old_comparison(tmp_path, monkeypatch, action):
    _, _, _, _, service, _, authenticated, payload, result = prepared(tmp_path, monkeypatch)
    draft = request(result)
    if action == 'expire':
        service.return_snapshot['created'] -= 301
    elif action == 'logout':
        authenticated[0] = False
    elif action == 'invalidate':
        service.invalidate()
    elif action == 'cancel':
        assert service.compare_return({**payload, 'approved_read': False})['result']['status'] == 'cancelled'
    else:
        assert service.compare_return(payload)['ok']
    reply = service.continue_return(draft)
    assert not reply['ok']
    assert reply['reason'] in {'professional_return_expired', 'not_authenticated', 'professional_return_target'}


def test_new_plan_conflict_and_uncertain_state_do_not_retry(tmp_path, monkeypatch):
    _, _, current, editor, service, _, _, payload, result = prepared(tmp_path, monkeypatch)
    task = service.continue_return(request(result))['result']
    plan = service.plan_group(task['task_id'])['result']
    count = len(list(editor.audit.iterdir()))
    current.write_bytes(current.read_bytes() + b' ')
    assert not service.confirm_group(task['task_id'], plan['id'], True)['ok']
    assert len(list(editor.audit.iterdir())) == count
    service.tasks[task['task_id']]['uncertain'] = True
    result = service.compare_return(payload)['result']
    assert service.continue_return(request(result))['reason'] == 'professional_task_unresolved'
    assert len(list(editor.audit.iterdir())) == count


def test_user_only_change_requires_ack_and_binds_current_without_replaying_old_text(tmp_path, monkeypatch):
    _, _, current, _, service, _, _, _, result = prepared(tmp_path, monkeypatch)
    payload = request(result, text='用户自行补充', replacement='用户明确补充')
    assert service.continue_return({**payload, 'acknowledge_current': False})['reason'] == 'professional_return_acknowledgement'
    reply = service.continue_return(payload)
    assert reply['ok'], reply
    task = reply['result']
    assert task['source_sha256'] == digest(current) and task['receipts'] == []
    assert task['items'][0]['before'] == '用户自行补充' and task['items'][0]['after'] == '用户明确补充'


def test_continuation_archive_reopens_without_old_comparison_or_approval(tmp_path, monkeypatch):
    _, _, current, _, service, _, _, _, result = prepared(tmp_path, monkeypatch)
    task = service.continue_return(request(result))['result']
    old_plan = service.plan_group(task['task_id'])['result']
    archived = service.plan_archive(task['task_id'])
    assert not archived['ok']
    assert service.confirm_group(task['task_id'], old_plan['id'], False)['result']['status'] == 'cancelled'
    archived = service.plan_archive(task['task_id'])
    assert archived['ok'], archived
    saved = service.confirm_archive(archived['result']['plan_id'], True)
    assert saved['ok'], saved
    editor = Editor(tmp_path, require_confirmation=True, render_validation=True)
    reopened = DocumentTasks(RLock(), lambda: (True, False, str(current)), lambda: editor, lambda: current, OutputDependencies())
    restored = reopened.restore(saved['result']['save_id'])
    assert restored['ok'], restored
    assert not restored['result']['write_approval_restored'] and not restored['result']['rollback_authority_restored']
    assert reopened.return_snapshot is None
    assert not reopened.continue_return(request(result))['ok']
    assert not reopened.confirm_group(task['task_id'], old_plan['id'], True)['ok']
    assert reopened.plan_group(task['task_id'])['ok']
