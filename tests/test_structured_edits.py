from copy import deepcopy
import json
import threading
from types import SimpleNamespace

from docx import Document
from lxml import etree as ET
import pytest

from office_agent.advanced_writes import _parts
from office_agent.grounded_edits import seal
from office_agent.professional_edits import _inventory, W
from office_agent.structured_edits import StructuredDocx, ParagraphInsertion, ReviewComment, anchor_record, transform, verify
from office_agent.writes import Editor, GuardError, digest
from office_agent.webui.bridge import Api
from test_professional_edits import fake_render


def fixture(tmp_path):
    path = tmp_path / 'report.docx'
    doc = Document()
    doc.add_paragraph('First ordinary anchor.')
    p = doc.add_paragraph()
    p.add_run('Measured latency ').bold = True
    p.add_run('20 ms in dry samples.').italic = True
    doc.add_paragraph('References')
    doc.add_table(rows=1, cols=1).cell(0, 0).text = 'Unchanged cell'
    doc.save(path)
    return path, Editor(tmp_path, require_confirmation=True, render_validation=True)


def request_for(path, insertions=None, comments=None):
    version = digest(path)
    _, rows = _inventory(_parts(path), version)
    insertions = insertions if insertions is not None else [ParagraphInsertion(anchor=rows[0]['locator_id'], side='after', text='Supplement one.'), ParagraphInsertion(anchor=rows[0]['locator_id'], side='after', text='Supplement two.')]
    comments = comments or []
    intent = {'source_sha256': version, 'anchors': {r['locator_id']: anchor_record(r, rows, version) for r in rows},
              'insertions': [r.model_dump() for r in insertions], 'comments': [r.model_dump() for r in comments],
              'date': '2026-10-18T00:00:00Z', 'catalog': {}, 'bindings': {}, 'authorized': [], 'bibliography': [], 'bibliography_anchor': None}
    return StructuredDocx(path=str(path), source_sha256=version, insertions=insertions, comments=comments, intent=intent, intent_seal=seal(intent))


def test_insert_order_non_target_preservation_copy_reopen_audit_undo(tmp_path, monkeypatch):
    path, editor = fixture(tmp_path)
    before = _parts(path)
    req = request_for(path)
    updated, _, after = transform(before, req)
    texts = [p.text for p in Document_from_parts(updated).paragraphs]
    assert texts[:4] == ['First ordinary anchor.', 'Supplement one.', 'Supplement two.', 'Measured latency 20 ms in dry samples.']
    verify(before, updated, req)
    monkeypatch.setattr('office_agent.advanced_writes.render_gate', fake_render)
    plan = editor.plan(req)
    assert not list(tmp_path.glob('*.agent-*.docx'))
    output = editor.execute(plan.id, confirmed=True)
    assert output['status'] == 'verified' and _parts(path) == before
    assert [p.text for p in Document(output['output']).paragraphs] == texts
    assert editor.rollback(plan.id, confirmed=True)['status'] == 'rolled_back'
    assert _parts(path) == before


def Document_from_parts(parts):
    from io import BytesIO
    from zipfile import ZipFile, ZIP_DEFLATED
    stream = BytesIO()
    with ZipFile(stream, 'w', ZIP_DEFLATED) as z:
        for k, v in parts.items():
            z.writestr(k, v)
    stream.seek(0)
    return Document(stream)


def test_comment_exact_range_and_explicit_parts_tamper_rejection(tmp_path):
    path, editor = fixture(tmp_path)
    _, rows = _inventory(_parts(path), digest(path))
    comment = ReviewComment(anchor=rows[1]['locator_id'], start=17, end=22, quote='20 ms', text='Check sample conditions before generalizing.')
    req = request_for(path, insertions=[], comments=[comment])
    before = _parts(path)
    after, _, receipt = transform(before, req)
    changed = {k for k in before.keys() | after.keys() if before.get(k) != after.get(k)}
    assert changed == {'word/document.xml', 'word/comments.xml', 'word/_rels/document.xml.rels', '[Content_Types].xml'}
    doc = Document_from_parts(after)
    assert [p.text for p in doc.paragraphs] == [p.text for p in Document(path).paragraphs]
    assert next(iter(doc.comments)).author == 'Office Agent'
    verify(before, after, req)
    tampered = dict(after)
    tampered['word/styles.xml'] += b' '
    with pytest.raises(GuardError, match='non_target_part'):
        verify(before, tampered, req)
    with pytest.raises(GuardError, match='existing_comments'):
        transform(after, req)


def test_duplicate_stale_anchor_and_source_refused(tmp_path):
    path, editor = fixture(tmp_path)
    req = request_for(path)
    bad = req.model_copy(update={'insertions': [req.insertions[0], req.insertions[0]]})
    bad.intent = {**bad.intent, 'insertions': [r.model_dump() for r in bad.insertions]}
    bad.intent_seal = seal(bad.intent)
    with pytest.raises(GuardError, match='overlapping'):
        transform(_parts(path), bad)
    plan = editor.plan(req)
    doc = Document(path)
    doc.paragraphs[0].text = 'Externally changed'
    doc.save(path)
    with pytest.raises(GuardError):
        editor.execute(plan.id, confirmed=True)
    assert not list(tmp_path.glob('*.agent-*.docx'))


def test_host_anchors_cancel_and_logout(tmp_path, monkeypatch):
    path, editor = fixture(tmp_path)
    authenticated = [True]
    app = SimpleNamespace(path=path, editor=editor, busy=False, window=None)
    auth = SimpleNamespace(root=tmp_path / 'auth', status=lambda: {'authenticated': authenticated[0]})
    api = Api(app, auth=auth, settings_root=tmp_path)
    source = api.read_structure_anchors()['result']
    payload = {'source_sha256': source['document']['source_sha256'], 'insertions': [{'anchor': source['anchors'][0]['locator_id'], 'side': 'after', 'text': 'Added paragraph.'}]}
    planned = api.plan_structure_edit(payload)
    assert planned['ok'], planned
    assert api.confirm_professional_edit(planned['result']['id'], False)['result']['status'] == 'cancelled'
    planned = api.plan_structure_edit(payload)
    authenticated[0] = False
    assert not api.confirm_professional_edit(planned['result']['id'], True)['ok']
    assert not list(tmp_path.glob('*.agent-*.docx'))


def test_no_placeholder_evidence_insertion_bibliography_and_refused_missing_scope(tmp_path, monkeypatch):
    from test_grounded import setup_service
    from office_agent.webui.structured import Structured
    from office_agent.webui.citations import Citations
    grounded, path, record, authenticated, busy = setup_service(tmp_path, monkeypatch)
    doc = Document()
    doc.add_paragraph('Experiment overview.')
    doc.add_paragraph('References')
    doc.save(path)
    capture = grounded.capture(record, 'paper.pdf')['result']
    service = Structured(grounded._lock, grounded._scope, grounded._editor, grounded._selected)
    service.grounded = grounded
    service.citations = Citations(grounded._lock, grounded._scope, grounded._editor, grounded._selected, grounded)
    data = service.read_structure()['result']
    original = digest(path)
    payload = {'source_sha256': original, 'insertions': [{'anchor': data['anchors'][0]['locator_id'], 'side': 'after',
                'text': 'The measured latency was 20 ms.', 'attributions': [{'capture_id': capture['capture_id'],
                    'evidence_ids': [capture['fragments'][0]['id']], 'relation': 'supports', 'applicability': 'Dry experiment',
                    'limitation': 'Self authored fixture only', 'assessment': 'user_reviewed', 'identity_reviewed': True}]}],
               'bibliography_anchor': data['anchors'][1]['locator_id']}
    plan = service.plan_structure(payload)
    assert plan['ok'], plan
    assert '[1]' in plan['result']['after'] and 'bibliography_insert' in plan['result']['after']
    planned_evidence = next(row['attributions'] for row in json.loads(plan['result']['after'])['items'] if row['operation'] == 'insert')
    assert planned_evidence[0]['evidence_fragments'][0]['id'] == capture['fragments'][0]['id']
    saved = service.confirm(plan['result']['id'], True)
    assert saved['ok'] and saved['result']['status'] == 'saved', saved
    output = Document(saved['result']['output'])
    assert output.paragraphs[1].text == 'The measured latency was 20 ms. [1]'
    assert output.paragraphs[2].text == 'References'
    assert output.paragraphs[3].text.startswith('[1] ') and '10.1234/expt.01' in output.paragraphs[3].text
    audit = json.loads((tmp_path / '.office_agent' / plan['result']['id'] / 'audit.json').read_bytes())
    persisted = next(row['attributions'] for row in json.loads(audit['after'])['items'] if row['operation'] == 'insert')
    assert persisted == planned_evidence
    assert persisted[0]['source_document']['source_sha256'] == capture['document']['source_sha256']
    assert digest(path) == original
    assert service.rollback(plan['result']['id'], True)['result']['status'] == 'rolled_back'
    payload['insertions'][0]['text'] = 'The measured latency was 99 ms.'
    assert service.plan_structure(payload)['reason'] == 'professional_unsupported_fact'


def test_claim_review_keeps_unreviewed_and_rejects_unknown_evidence(tmp_path, monkeypatch):
    from test_grounded import setup_service
    from office_agent.webui.structured import Structured
    grounded, path, record, authenticated, busy = setup_service(tmp_path, monkeypatch)
    capture = grounded.capture(record, 'paper.pdf')['result']
    service = Structured(grounded._lock, grounded._scope, grounded._editor, grounded._selected)
    service.grounded = grounded
    locator = service.read_structure()['result']['anchors'][0]['locator_id']
    evidence = capture['fragments'][0]['id']
    class Model:
        def suggest(self, payload):
            return {'cards': [{'claim': 'Latency was measured.', 'evidence_ids': [evidence], 'relation': 'supports',
                               'applicability': 'dry samples', 'limitation': 'partial evidence', 'conflicts': []}],
                    'suggestions': [], 'missing_evidence': [], 'stop_reason': 'bounded review complete'}
    service.model_factory = Model
    reply = service.review_claim(locator, [{'capture_id': capture['capture_id'], 'evidence_ids': [evidence]}])
    assert reply['ok'] and reply['result']['user_reviewed'] is False and reply['result']['model_calls'] == 1
    assert not list(tmp_path.glob('*.agent-*.docx'))
    assert not service.review_claim(locator, [{'capture_id': capture['capture_id'], 'evidence_ids': ['ev:' + '0' * 64]}])['ok']


def test_grouped_clean_replacement_insert_and_comment_overlap_guard(tmp_path):
    from office_agent.professional_edits import LocalChange
    path, editor = fixture(tmp_path)
    req = request_for(path)
    _, rows = _inventory(_parts(path), digest(path))
    replacement = LocalChange(locator_id=rows[0]['locator_id'], expected_text_sha256=rows[0]['text_sha256'],
                              new_text='First ordinary anchor clarified.', reason='Clarify an existing sentence.')
    req.replacements = [replacement]
    req.intent['replacements'] = [replacement.model_dump()]
    req.intent_seal = seal(req.intent)
    after, before_items, after_items = transform(_parts(path), req)
    assert Document_from_parts(after).paragraphs[0].text == replacement.new_text
    assert any(r['operation'] == 'replace' for r in after_items)
    verify(_parts(path), after, req)
    req.comments = [ReviewComment(anchor=rows[0]['locator_id'], start=0, end=5, quote='First', text='Overlapping review')]
    req.intent['comments'] = [r.model_dump() for r in req.comments]
    req.intent_seal = seal(req.intent)
    with pytest.raises(GuardError, match='overlapping'):
        transform(_parts(path), req)


def test_inherited_numbered_style_is_not_copied(tmp_path):
    from docx.enum.style import WD_STYLE_TYPE
    from docx.oxml import OxmlElement
    path, editor = fixture(tmp_path)
    doc = Document(path)
    parent = doc.styles.add_style('UnsafeNumbered', WD_STYLE_TYPE.PARAGRAPH)
    parent.element.get_or_add_pPr().append(OxmlElement('w:numPr'))
    child = doc.styles.add_style('InheritedNumbered', WD_STYLE_TYPE.PARAGRAPH)
    child.base_style = parent
    doc.paragraphs[0].style = child
    doc.save(path)
    req = request_for(path)
    with pytest.raises(GuardError, match='structure_style'):
        transform(_parts(path), req)


def test_insertion_coordinates_existing_numbers_and_duplicate_doi_as_group(tmp_path, monkeypatch):
    from test_grounded import setup_service
    from test_literature import crossref
    from office_agent.literature import normalize
    from office_agent.webui.structured import Structured
    from office_agent.webui.citations import Citations
    grounded, path, first, auth, busy = setup_service(tmp_path, monkeypatch)
    second = normalize(crossref('Independent experiment', '10.1234/EXPT.02'), 'crossref')
    grounded.literature.records[second['id']] = second
    grounded.literature.decisions[second['id']] = {'decision': 'include', 'reason': 'Owned fixture'}
    doc = Document()
    for text in ['Existing result [4].', 'Another result [7].', 'References',
                 '[4] Second experiment doi: 10.1234/expt.02', '[7] First experiment doi: 10.1234/expt.01']:
        doc.add_paragraph(text)
    doc.save(path)
    captured = grounded.capture(first, 'paper.pdf')['result']
    service = Structured(grounded._lock, grounded._scope, grounded._editor, grounded._selected)
    service.grounded = grounded
    service.citations = Citations(grounded._lock, grounded._scope, grounded._editor, grounded._selected, grounded)
    data = service.read_structure()['result']
    anchors = data['anchors']
    payload = {'source_sha256': digest(path), 'insertions': [{'anchor': anchors[0]['locator_id'], 'side': 'before',
                'text': 'The measured latency was 20 ms.', 'attributions': [{'capture_id': captured['capture_id'],
                'evidence_ids': [captured['fragments'][0]['id']], 'relation': 'supports', 'applicability': 'Owned experiment',
                'limitation': 'Development fixture only', 'assessment': 'user_reviewed', 'identity_reviewed': True}]}],
               'bibliography_anchor': anchors[2]['locator_id'], 'bibliography': [r['locator_id'] for r in anchors[3:]],
               'authorized': [r['locator_id'] for r in anchors[1:]], 'number_records': {'4': second['id'], '7': first}}
    assert service.plan_structure(payload)['reason'] == 'professional_citation_scope'
    assert not list(tmp_path.glob('*.agent-*.docx'))
    payload['authorized'].append(anchors[0]['locator_id'])
    plan = service.plan_structure(payload)
    assert plan['ok'], plan
    saved = service.confirm(plan['result']['id'], True)
    assert saved['result']['status'] == 'saved', saved
    paragraphs = [p.text for p in Document(saved['result']['output']).paragraphs]
    assert paragraphs[:3] == ['The measured latency was 20 ms. [1]', 'Existing result [2].', 'Another result [1].']
    assert sum('10.1234/expt.01' in text for text in paragraphs) == 1
    assert sum('10.1234/expt.02' in text for text in paragraphs) == 1
    assert service.rollback(plan['result']['id'], True)['result']['status'] == 'rolled_back'
