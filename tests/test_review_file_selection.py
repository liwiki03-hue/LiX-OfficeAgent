from copy import deepcopy
import shutil
from types import SimpleNamespace

import pytest

from office_agent import render_handoff as handoff
from office_agent.webui.bridge import Api
from office_agent.writes import Editor
from tests.test_render_handoff import prepared


@pytest.mark.parametrize('different_root', [False, True])
def test_file_selection_keeps_only_same_root_artifacts_and_never_approval(tmp_path, different_root):
    from docx import Document
    source = tmp_path / 'source.docx'
    doc = Document()
    doc.add_paragraph('选择样例')
    doc.save(source)
    root = tmp_path / 'different' if different_root else tmp_path
    root.mkdir(exist_ok=True)
    selected = root / 'selected.docx'
    shutil.copyfile(source, selected)
    state = [True]
    auth = SimpleNamespace(root=tmp_path / 'auth', status=lambda: {'authenticated': state[0]})
    app = SimpleNamespace(path=source, editor=Editor(tmp_path, require_confirmation=True), busy=False,
                          window=None, _build_model=lambda: None, _clear_private_state=lambda: None)
    api = Api(app, auth=auth, settings_root=tmp_path)
    def choose():
        app.editor = Editor(root, require_confirmation=True)
        app.path = app.editor.space.file(str(selected))
        return {'ok': True}
    app.select_file = choose
    handoff.clear()
    try:
        report, pdfs = prepared(tmp_path)
        second = {**deepcopy(report), 'output_sha256': 'second'}
        assert handoff.offer(tmp_path, report, pdfs)
        assert handoff.offer(tmp_path, second, pdfs)
        api._range_review.reviews['old'] = {'consent': 'old grant'}
        api._professional._saved['old'] = {'authority': 'old grant'}
        assert api.select_file()['ok']
        assert not api._range_review.reviews and not api._professional._saved
        assert len(handoff._entries()) == (0 if different_root else 2)
        assert not api.range_review_page('old', 0, 'before')['ok']
        api._clear_private_state()
        assert not handoff._entries()
    finally:
        api._clear_private_state()
        handoff.clear()
