import json
from pathlib import Path
from threading import RLock

from docx import Document
import pytest

from office_agent.range_alignment import locate, align, anchors_for
from office_agent.range_edits import view
from office_agent.output_dependencies import OutputDependencies
from office_agent.webui.professional import Professional
from office_agent.webui.range_review import RangeReview
from office_agent.writes import Editor, digest
from tests.test_range_edits import operation, request
from tests.test_range_fields import field_sample
from tests.test_professional_edits import fake_render


def test_context_disambiguates_occurrences_and_adjacent_pages():
    pages = [{'text': '其他条件 共同结论 后文。独特条件 共同'}, {'text': '结论 必须复核。'}]
    assert locate(pages, {'text': '共同结论', 'prefix': '独特条件', 'suffix': '必须复核'}) == {'pages': [1, 2], 'status': 'unique', 'matches': 1}
    assert locate(pages, {'text': '共同结论'})['status'] == 'ambiguous'
    assert locate([{'text': '重复结论 重复结论'}], {'text': '重复结论'})['status'] == 'ambiguous'
    assert locate(pages, {'text': '', 'prefix': '', 'suffix': ''})['status'] == 'not_located'
    assert locate([{'text': 'Cafe\u0301实验'}], {'text': 'Café实验'})['status'] == 'unique'
    assert locate([{'text': 'therapist'}], {'text': 'the rapist'})['status'] == 'not_located'
    assert locate([{'text': 'alpha'}, {'text': 'beta'}], {'text': 'alphabeta'})['status'] == 'not_located'
    assert locate([{'text': 'alpha'}, {'text': 'beta'}], {'text': 'alpha beta'})['pages'] == [1, 2]
    assert locate([{'text': 'alpha'}, {'text': ''}, {'text': 'beta'}], {'text': 'alpha beta'})['status'] == 'not_located'
    assert locate([{'text': '甲乙'}, {'text': '丙丁'}, {'text': '戊己'}], {'text': '甲乙丙丁戊己'})['status'] == 'not_located'


def test_impact_and_structural_action_availability(field_sample):
    from office_agent.advanced_writes import _parts
    from office_agent.range_edits import transform
    path, editor, data = field_sample
    target, nearby = data['items'][:2]
    assert target['action_support']['clean']['insert_before'] == 'professional_structure_scope'
    assert target['action_support']['clean']['replace'] == 'professional_reference_target_locked'
    req = request(field_sample, [operation(nearby, 0, 3, '研究结论')])
    _, _, after = transform(_parts(path, range_edit=True), req)
    impact = after[0]['verification']['reference_impact']
    assert impact['needs_review_total'] == 2
    assert impact['items'][0]['target_location'] == target['location']
    assert 'reference_surroundings' in impact['items'][0]['reasons']
    assert 'possible_pagination' in impact['items'][1]['reasons']
    assert not impact['dynamic_word_pagination_verified'] and not impact['complete_dependency_graph']


def prepared(tmp_path, monkeypatch):
    source = tmp_path / 'source.docx'
    doc = Document()
    doc.add_paragraph('独特条件 原结论 仅适用干燥环境。')
    doc.save(source)
    editor = Editor(tmp_path, require_confirmation=True, render_validation=True)
    state = [True]
    lock = RLock()
    scope = lambda: (state[0], False, str(source))
    service = Professional(lock, scope, lambda: editor, lambda: source)
    dependencies = OutputDependencies()
    service.dependencies = dependencies
    def render(*args, **kwargs):
        return {**fake_render(*args, **kwargs), 'runtime_sha256': 'a' * 64,
                'render': {'status': 'different', 'anchor_alignment': [{'location': 'loc:' + 'b' * 64, 'reason': '精确措辞', 'action': 'replace',
                     'before_pages': [1], 'after_pages': [1], 'before_status': 'unique', 'after_status': 'unique'}]}}
    monkeypatch.setattr('office_agent.advanced_writes.render_gate', render)
    data = service.read_ranges()['result']
    planned = service.plan_ranges({'source_sha256': digest(source), 'operations': [operation(data['items'][0], 5, 8, '新结论')]})
    assert planned['ok'], planned
    saved = service.confirm(planned['result']['id'], True)
    assert saved['ok'], saved
    receipt = saved['result']['save_id']
    review = RangeReview(lock, scope, lambda: editor, lambda: source, dependencies, service.rollback)
    return review, service, editor, source, receipt, state


def test_review_archive_restore_revalidates_and_never_restores_grants(tmp_path, monkeypatch):
    review, writer, editor, source, receipt, state = prepared(tmp_path, monkeypatch)
    identifier = review.start(receipt, [])['result']['review_id']
    assert review.open(identifier, False)['result']['status'] == 'cancelled'
    assert review.page(identifier, 0, 'before')['reason'] == 'professional_review_read_consent'
    opened = review.open(identifier, True)
    assert opened['ok'], opened
    assert review.mark(identifier, 0, 'ignored', '')['ok'] is False
    marked = review.mark(identifier, 0, 'needs_review', '核对分页')
    assert marked['ok'], marked
    plan = review.plan_archive(identifier, marked['result']['revision'])
    assert plan['ok'], plan
    assert review.confirm_archive(plan['result']['plan_id'], False)['result']['status'] == 'rejected'
    plan = review.plan_archive(identifier, marked['result']['revision'])
    saved = review.confirm_archive(plan['result']['plan_id'], True)
    assert saved['ok'], saved
    archive = saved['result']['save_id']
    writer.invalidate()
    review.invalidate()
    restored = review.restore_archive(archive)
    assert restored['ok'], restored
    assert not restored['result']['restores_approval'] and not restored['result']['restores_rollback_authority']
    identifier = restored['result']['review_id']
    assert review.page(identifier, 0, 'before')['reason'] == 'professional_review_read_consent'
    assert review.open(identifier, True)['result']['items'][0]['state'] == 'needs_review'
    assert not writer._saved and not review._saved
    assert not review.page(identifier, True, 'before')['ok']
    assert not review.page(identifier, 0, '../source.docx')['ok']
    removal = review.plan_remove_archive(archive)
    assert removal['ok'], removal
    cancelled = review.confirm_remove_archive(removal['result']['rollback_id'], False)
    assert cancelled['ok'] and cancelled['result']['status'] == 'rejected'
    removal = review.plan_remove_archive(archive)
    removed = review.confirm_remove_archive(removal['result']['rollback_id'], True)
    assert removed['ok'] and removed['result']['status'] == 'rolled_back'
    assert source.exists()
    source.write_bytes(source.read_bytes() + b'\n')
    assert review.open(identifier, True)['reason'] == 'professional_review_changed'
    assert review.mark(identifier, 0, 'viewed')['reason'] == 'professional_review_changed'
    state[0] = False
    assert review.open(identifier, True)['reason'] == 'not_authenticated'
    assert not review.reviews and review.cache is None


def test_review_rejects_stale_item_plan_audit_and_output_changes(tmp_path, monkeypatch):
    review, writer, editor, source, receipt, state = prepared(tmp_path, monkeypatch)
    identifier = review.start(receipt, [])['result']['review_id']
    review.open(identifier, True)
    plan = review.plan_archive(identifier, 1)['result']['plan_id']
    assert review.mark(identifier, 0, 'viewed')['ok']
    assert review.confirm_archive(plan, True)['reason'] == 'professional_review_changed'
    audit_path = editor.audit / receipt / 'audit.json'
    audit = json.loads(audit_path.read_bytes())
    output = Path(audit['output'])
    output.write_bytes(output.read_bytes() + b'\n')
    assert review.page(identifier, 0, 'after')['reason'] == 'professional_review_changed'
    assert review.cache is None
