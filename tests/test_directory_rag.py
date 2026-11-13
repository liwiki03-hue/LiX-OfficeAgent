import json
import subprocess

from docx import Document
import pytest

from office_agent.rag import RagError
from office_agent.rag_workspace import WorkspaceRag
from office_agent.writes import Workspace, digest
from test_rag_api import make
from test_rag_ui import NODE, PROBE, STATIC


SESSION = 'a' * 32


def document(path, text):
    doc = Document()
    doc.add_paragraph(text)
    doc.save(path)
    return path


class SyntheticEmbedder:
    identity = {'channel': 'local', 'model': 'cross-document-test-model', 'dimension': 2}

    def __init__(self):
        self.calls = []

    def embed(self, texts):
        self.calls.append(texts)
        return [[1., 0.] if 'vehicle' in text or 'car' in text else [0., 1.] for text in texts]


def test_directory_semantic_query_saves_only_after_independent_confirmation(tmp_path):
    first = document(tmp_path / 'vehicle.docx', 'vehicle registration')
    second = document(tmp_path / 'fruit.docx', 'apple orange')
    before = {path: digest(path) for path in (first, second)}
    embedder = SyntheticEmbedder()
    service = WorkspaceRag(Workspace(tmp_path), embedder)
    plan = service.plan_directory_query(first.name, 'car', SESSION, top_k=1)
    assert plan['hit_count'] == 1 and 'hits' not in plan
    assert not list(tmp_path.glob('.office-rag-evidence-*'))
    assert service.snapshot()['exists'] is False
    assert len(embedder.calls) == 2
    denied = service.confirm_query(plan['plan_id'], SESSION, False)
    assert denied['wrote_nothing'] and denied['applied'] is False
    plan = service.plan_directory_query(first.name, 'car', SESSION, top_k=1)
    saved = service.confirm_query(plan['plan_id'], SESSION, True)
    assert saved['verified'] and saved['hits'][0]['source'] == 'vehicle.docx'
    assert saved['hits'][0]['excerpt'] == 'vehicle registration'
    assert 'text' not in saved['hits'][0]
    assert {path: digest(path) for path in before} == before
    payload = json.loads((tmp_path / saved['evidence']).read_text(encoding='utf-8'))
    assert payload['hits'][0]['text'] == 'vehicle registration'


@pytest.mark.parametrize('question,top_k', [('', 5), ('x' * 201, 5), ('car', 11), ('car', True)])
def test_invalid_directory_query_does_not_call_provider(tmp_path, question, top_k):
    selected = document(tmp_path / 'vehicle.docx', 'vehicle registration')
    embedder = SyntheticEmbedder()
    service = WorkspaceRag(Workspace(tmp_path), embedder)
    with pytest.raises(RagError):
        service.plan_directory_query(selected.name, question, SESSION, top_k=top_k)
    assert embedder.calls == []


def test_directory_evidence_source_changed_before_save_is_refused(tmp_path):
    selected = document(tmp_path / 'vehicle.docx', 'vehicle registration')
    service = WorkspaceRag(Workspace(tmp_path), SyntheticEmbedder())
    plan = service.plan_directory_query(selected.name, 'car', SESSION)
    document(selected, 'source changed')
    result = service.confirm_query(plan['plan_id'], SESSION, True)
    assert result['wrote_nothing'] and not result['applied']
    assert not list(tmp_path.glob('.office-rag-evidence-*'))


def test_authenticated_directory_api_rejects_logout_busy_and_client_source_paths(tmp_path):
    api, app, auth, root, selected, embedder, factories = make(tmp_path)
    source = document(root / 'one.docx', '预算成本安排')
    document(root / 'two.docx', '预算开支说明')
    app.path = source
    plan = api.plan_rag_directory_query('预算', 2)
    assert plan['ok'] and plan['plan']['hit_count'] == 2
    assert api.confirm_rag_query(plan['plan']['plan_id'], False)['result']['wrote_nothing']
    calls = len(embedder.calls)
    app.busy = True
    assert api.plan_rag_directory_query('预算')['reason'] == 'busy'
    app.busy = False
    auth.authenticated = False
    assert api.plan_rag_directory_query('预算')['reason'] == 'not_authenticated'
    assert len(embedder.calls) == calls


@pytest.mark.parametrize('scenario', ['directory', 'directory_denied', 'directory_late', 'directory_partial'])
def test_directory_dom_scope_permission_and_late_response(scenario):
    assert NODE is not None
    probe = PROBE.replace("nodes['rag-topk'].value='5';", "nodes['rag-directory']=new Element('rag-directory'); nodes['rag-topk'].value='5';", 1)
    probe = probe.replace('const raw=Object.fromEntries', "const dqp=()=>{const p=qp();p.plan.directory={document_count:1,skipped_document_count:0,corpus_truncated:false,coverage_complete:true,inventory:[{source:'report.docx',status:'indexed_in_memory',reason:null},{source:'old.doc',status:'skipped',reason:'legacy_format'}]};return p;}; methods.plan_rag_directory_query=async()=>dqp();\nconst raw=Object.fromEntries")
    probe = probe.replace('return p;}; methods.plan_rag_directory_query', "if(scenario==='directory_partial'){p.plan.directory.skipped_document_count=1;p.plan.directory.coverage_complete=false;p.plan.directory.corpus_truncated=true;p.plan.directory.inventory.push({source:'large.docx',status:'skipped',reason:'capacity'});}return p;}; methods.plan_rag_directory_query")
    marker = "if(scenario==='initial_auth')authenticated=false;"
    assert marker in probe
    block = """
if(scenario.startsWith('directory')) {
  fill(); nodes['rag-directory'].checked=true;
  if(scenario==='directory_denied') {
    modal=false; await click('query'); assert.equal(count('plan_rag_directory_query'),0);
  } else if(scenario==='directory_late') {
    const d=deferred(); methods.plan_rag_directory_query=()=>d.promise;
    const pending=click('query'); await tick(); await nodes['rag-directory'].fire('change');
    d.resolve(dqp()); await pending; assert.equal(count('confirm_rag_query'),0);
  } else {
    await click('query'); assert.equal(count('plan_rag_directory_query'),1);
    assert.equal(count('plan_rag_query'),0); assert.equal(count('modal'),2);
    assert.equal(count('confirm_rag_query'),1);
    assert(nodes['rag-hits'].textContent.includes('<img onerror=x>'));
    assert(nodes['rag-snapshot'].textContent.includes('old.doc'));
    if(scenario==='directory_partial')assert(nodes['rag-snapshot'].textContent.includes('覆盖不完整'));
  }
} else
"""
    anchor = "if(scenario==='snapshot'){"
    assert anchor in probe
    probe = probe.replace(anchor, block + anchor, 1)
    result = subprocess.run([NODE, '-e', probe, str(STATIC / 'rag.js'), scenario], capture_output=True,
                            text=True, encoding='utf-8', timeout=15)
    assert result.returncode == 0, result.stderr
    assert result.stdout == 'RAG_DOM_OK'
