from pathlib import Path
import json

from office_agent.range_alignment import locate, anchors_for, align
from office_agent.task_routing import candidate
from tests.test_task_routing import prepared


def test_host_paragraph_fallback_recovers_registered_real_render_failure(tmp_path, monkeypatch):
    source, selected, editor, service, task, auth = prepared(tmp_path, monkeypatch)
    details, native = candidate(editor, source, service.tasks[task['task_id']])
    assert native and len(native['render_anchors']) == 12
    root = Path(__file__).resolve().parents[1]
    snapshots = json.loads((root / 'deliverables/long-review-20261025/baseline1/RENDER_TEXT.json').read_text(encoding='utf8'))
    result = align(native['render_anchors'], snapshots)
    assert all(row['before_status'] == row['after_status'] == 'unique' for row in result)
    assert result[7]['before_method'] == result[7]['after_method'] == 'host-paragraph-context-fallback'
    assert result[7]['before_pages'] == result[7]['after_pages'] == [2]


def test_fallback_remains_exact_unique_and_does_not_override_ambiguity():
    context = {'text': '共同结论', 'prefix': '原段标题', 'suffix': '原段结束 下一段标题',
               'paragraph': {'prefix': '原段标题', 'suffix': '原段结束'}}
    pages = [{'text': '原段标题共同结论原段结束 页脚'}, {'text': '页眉 下一段标题'}]
    assert locate(pages, context)['pages'] == [1]
    assert locate(pages, context)['method'] == 'host-paragraph-context-fallback'
    assert locate([{'text': '原段标题共同结论原段结束 原段标题共同结论原段结束'}], context)['status'] == 'not_located'
    assert locate([{'text': 'therapist'}], {'text': 'the rapist', 'paragraph': {'prefix': '', 'suffix': ''}})['status'] == 'not_located'
    assert locate([{'text': '甲乙'}, {'text': ''}, {'text': '丙丁'}], {'text': '甲乙丙丁', 'paragraph': {}})['status'] == 'not_located'
    assert locate([{'text': '共同结论共同结论'}], {'text': '共同结论', 'paragraph': {'prefix': '不存在', 'suffix': ''}})['status'] == 'ambiguous'
    assert locate([{'text': '旧术语'}], {'text': '旧术语', 'paragraph': {}})['status'] == 'not_located'
