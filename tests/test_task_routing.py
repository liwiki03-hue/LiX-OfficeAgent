import json

from docx import Document
import pytest

from office_agent.document_tasks import natural_spec
from office_agent.writes import digest
from tests.test_document_tasks import prepared


def save_next(service, identifier):
    plan = service.plan_group(identifier, 'auto')
    assert plan['ok'], plan
    result = service.confirm_group(identifier, plan['result']['id'], True)
    assert result['ok'] and result['result']['status'] == 'saved', result
    return plan['result'], result['result']['task']


def test_native_route_complete_ledger_two_transactions_cancel_archive_undo(tmp_path, monkeypatch):
    source, selected, editor, service, task, auth = prepared(tmp_path, monkeypatch)
    before = digest(source)
    plan = service.plan_group(task['task_id'], 'auto')
    assert plan['ok'] and plan['result']['kind'] == 'word_batch_replace', plan
    assert plan['result']['affected_ledger_items'] == 30
    assert plan['result']['routing']['page_alignment_items'] == 12
    assert not plan['result']['routing']['page_alignment_complete']
    assert not editor.audit.exists()
    assert service.confirm_group(task['task_id'], plan['result']['id'], False)['result']['status'] == 'cancelled'
    assert not editor.audit.exists()
    first, task = save_next(service, task['task_id'])
    assert task['states'] == {'completed': 30, 'pending': 1} and task['verification_checks'][0]['remaining'] == 1
    current = tmp_path / task['current_name']
    archive = service.plan_archive(task['task_id'])
    saved = service.confirm_archive(archive['result']['plan_id'], False)
    assert saved['result']['status'] == 'rejected'
    second, task = save_next(service, task['task_id'])
    assert second['kind'] == 'word_table_cell_replace' and task['deterministic_complete']
    assert task['states'] == {'completed': 31} and len(task['receipts']) == 2
    output = tmp_path / task['current_name']
    assert Document(output).tables[0].cell(1, 0).text == '规范术语'
    audit = json.loads((editor.audit / second['id'] / 'audit.json').read_text(encoding='utf8'))
    assert audit['task_route_verification']['status'] == 'passed'
    for _ in range(2):
        assert service.undo_last(task['task_id'], True)['result']['status'] == 'rolled_back'
    assert not current.exists() and not output.exists() and digest(source) == before


def test_native_route_restore_current_replans_without_old_approval(tmp_path, monkeypatch):
    source, selected, editor, service, task, auth = prepared(tmp_path, monkeypatch)
    _, task = save_next(service, task['task_id'])
    archive = service.plan_archive(task['task_id'])
    saved = service.confirm_archive(archive['result']['plan_id'], True)
    selected[0] = tmp_path / task['current_name']
    service.invalidate()
    restored = service.restore(saved['result']['save_id'])
    assert restored['ok'] and not restored['result']['write_approval_restored']
    assert not service.undo_last(task['task_id'], True)['ok']
    _, task = save_next(service, task['task_id'])
    assert task['deterministic_complete']


@pytest.mark.parametrize('alteration', ['tracked', 'selected', 'excluded', 'mixed_format'])
def test_non_equivalent_semantics_keep_range_route(tmp_path, monkeypatch, alteration):
    source, selected, editor, service, task, auth = prepared(tmp_path, monkeypatch)
    if alteration == 'mixed_format':
        doc = Document(source)
        doc.paragraphs[1].clear()
        doc.paragraphs[1].add_run('旧').bold = True
        doc.paragraphs[1].add_run('术语').italic = True
        doc.save(source)
    spec = natural_spec(editor, source, '将全文“旧术语”替换为“规范术语”').model_dump()
    if alteration == 'tracked':
        spec['mode'] = 'tracked'
    if alteration == 'selected':
        row = service.navigate({'query': '旧术语'})['result']['items'][0]
        spec['requirements'][0].update(scope='selected', locators=[row['locator_id']])
    task = service.create(spec)['result']
    if alteration == 'excluded':
        body = next(item for item in service.tasks[task['task_id']]['items'] if '/table/' not in item['location'])
        assert service.exclude(task['task_id'], body['id'], '用户保留该项')['ok']
    plan = service.plan_group(task['task_id'], 'auto')
    assert plan['ok'] and plan['result']['kind'] == 'range_docx', plan
    assert not editor.audit.exists()


def test_route_old_plan_invalid_after_comparison_source_change_and_logout(tmp_path, monkeypatch):
    source, selected, editor, service, task, auth = prepared(tmp_path, monkeypatch)
    native = service.plan_group(task['task_id'], 'auto')['result']
    ranges = service.plan_group(task['task_id'], 'ranges')['result']
    assert native['id'] not in editor.plans
    assert not service.confirm_group(task['task_id'], native['id'], True)['ok']
    assert not service.confirm_group(task['task_id'], ranges['id'], True)['ok']
    plan = service.plan_group(task['task_id'], 'auto')['result']
    doc = Document(source)
    doc.add_paragraph('外部修改')
    doc.save(source)
    assert service.confirm_group(task['task_id'], plan['id'], True)['reason'] == 'professional_source_changed'
    auth[0] = False
    assert service.plan_group(task['task_id'], 'auto')['reason'] == 'not_authenticated'
    assert not editor.audit.exists()


def test_routed_unexpected_change_refused_before_publication(tmp_path, monkeypatch):
    source, selected, editor, service, task, auth = prepared(tmp_path, monkeypatch)
    plan = service.plan_group(task['task_id'], 'auto')['result']
    from office_agent.task_routing import verify_package
    def corrupted(path, expected):
        doc = Document(path)
        doc.paragraphs[0].text = '意外变更'
        doc.save(path)
        verify_package(path, expected)
    monkeypatch.setattr('office_agent.task_routing.verify_package', corrupted)
    result = service.confirm_group(task['task_id'], plan['id'], True)
    assert result['result']['status'] == 'uncertain'
    assert list(tmp_path.glob('*.docx')) == [source]
    assert service.plan_group(task['task_id'], 'auto')['reason'] == 'professional_task_unresolved'


def test_native_plan_expiration_and_source_identity_guard(tmp_path, monkeypatch):
    source, selected, editor, service, task, auth = prepared(tmp_path, monkeypatch)
    plan = service.plan_group(task['task_id'], 'auto')['result']
    writer = service.writers[task['task_id']]
    writer._pending[plan['id']]['created'] -= 301
    assert service.confirm_group(task['task_id'], plan['id'], True)['reason'] == 'professional_plan_expired'
    assert plan['id'] not in editor.plans and not editor.audit.exists()


def test_distributed_cross_run_format_is_not_replaced_by_start_run_policy(tmp_path, monkeypatch):
    source, selected, editor, service, task, auth = prepared(tmp_path, monkeypatch)
    doc = Document(source)
    doc.paragraphs[2].clear()
    doc.paragraphs[2].add_run('旧').bold = True
    doc.paragraphs[2].add_run('术语').italic = True
    doc.save(source)
    task = service.create('将全文“旧术语”替换为“新术语”')['result']
    plan = service.plan_group(task['task_id'], 'auto')
    assert plan['ok'] and plan['result']['kind'] == 'range_docx', plan
    assert not editor.audit.exists()


@pytest.mark.parametrize('body_count,cell_count', [(1, 1), (0, 3)])
def test_routing_does_not_turn_combined_task_into_more_transactions(tmp_path, monkeypatch, body_count, cell_count):
    source, selected, editor, service, task, auth = prepared(tmp_path, monkeypatch)
    doc = Document()
    for _ in range(body_count):
        doc.add_paragraph('旧术语')
    table = doc.add_table(rows=1, cols=cell_count)
    for cell in table.rows[0].cells:
        cell.text = '旧术语'
    doc.save(source)
    task = service.create('将全文“旧术语”替换为“规范术语”')['result']
    plan = service.plan_group(task['task_id'], 'auto')
    assert plan['ok'] and plan['result']['kind'] == 'range_docx', plan
    assert not editor.audit.exists()


def test_native_anchors_use_all_approved_neighbors_and_prior_length_changes(tmp_path, monkeypatch):
    source, selected, editor, service, task, auth = prepared(tmp_path, monkeypatch)
    from office_agent.task_routing import candidate
    details, native = candidate(editor, source, service.tasks[task['task_id']])
    assert details['selected'] == 'word_batch_replace'
    assert all('旧术语' not in json.dumps(row['contexts']['after'], ensure_ascii=False) for row in native['render_anchors'])
    doc = Document()
    doc.add_paragraph('旧术语；' * 13)
    doc.save(source)
    task = service.create('将全文“旧术语”替换为“规范术语”')['result']
    details, native = candidate(editor, source, service.tasks[task['task_id']])
    assert native is not None and details['covered_items'] == 13
    assert len(native['render_anchors']) == 12
    assert all(row['contexts']['after']['text'] == '规范术语' for row in native['render_anchors'])


def test_routed_delivery_supports_read_only_return_without_restoring_rights(tmp_path, monkeypatch):
    source, selected, editor, service, task, auth = prepared(tmp_path, monkeypatch)
    original = digest(source)
    for _ in range(2):
        _, task = save_next(service, task['task_id'])
    delivered = tmp_path / task['current_name']
    doc = Document(delivered)
    doc.paragraphs[0].text = '用户重新另存的报告'
    returned = tmp_path / 'returned.docx'
    doc.save(returned)
    hashes = [digest(path) for path in (source, delivered, returned)]
    selected[0] = returned
    service.invalidate()
    result = service.compare_return({'baseline_name': source.name, 'delivered_name': delivered.name,
        'receipts': task['receipts'], 'approved_read': True})
    assert result['ok'], result
    assert result['result']['classifications'].get('system_suggestion_still_present', 0) > 0
    assert hashes == [digest(path) for path in (source, delivered, returned)] and digest(source) == original
    assert not service.confirm_group(task['task_id'], task['receipts'][-1], True)['ok']


@pytest.mark.parametrize('field', ['acceptance', 'goal', 'request'])
def test_extra_free_text_obligations_cannot_be_silently_routed_to_partial_page_coverage(tmp_path, monkeypatch, field):
    source, selected, editor, service, task, auth = prepared(tmp_path, monkeypatch)
    spec = natural_spec(editor, source, '将全文“旧术语”替换为“规范术语”').model_dump()
    if field == 'request':
        spec['request'] += '，并对每个目标逐项提供页图核对'
    else:
        spec['requirements'][0][field] = '每个目标都必须有独立完整页图审阅'
    task = service.create(spec)['result']
    plan = service.plan_group(task['task_id'], 'auto')
    assert plan['ok'] and plan['result']['kind'] == 'range_docx', plan
    assert '验收条件' in plan['result']['routing']['reason']
    assert not editor.audit.exists()
