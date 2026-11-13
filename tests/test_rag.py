"""Protocol tests: not a claim of real embedding quality or SQLite integration."""
import pytest

from office_agent.rag import RagError, RagService, structure_chunks


def block(text='正文', **overrides):
    return {'source':'report.docx','location':'paragraph:1','type':'paragraph','text':text,'group':'body',**overrides}


class Embedder:
    identity = {'channel':'local','model':'fixture-semantic-only','dimension':2}
    def __init__(self):
        self.calls = []
        self.result = None
        self.on_embed = lambda: None
    def embed(self, texts):
        self.calls.append(texts)
        self.on_embed()
        return self.result if self.result is not None else [[1.0,0.0] for _ in texts]


class Store:
    def __init__(self):
        self.calls = []
        self.metadata = {'revision':1,'model':'fixture-semantic-only','dimension':2,'uncertain':False}
        self.failure = False
        self.result = {'status':'saved','applied':True,'save_id':'saved'}
        self.hits = [{'doc_id':'document','source_fingerprint':'a'*64,'text':'文'*500,'source':'report.docx','location':'paragraph:1'}]
    def snapshot(self):
        return self.metadata
    def invalidate(self):
        self.calls.append(('invalidate',))
    def plan_replace(self, *args):
        self.calls.append(('plan_replace',*args))
        return {'plan_id':'inner'}
    def confirm_replace(self, *args):
        self.calls.append(('confirm_replace',*args))
        if self.failure:
            raise RuntimeError('fixture-private-error')
        return self.result
    def query(self, *args, **kwargs):
        self.calls.append(('query',args,kwargs))
        return self.hits


def setup():
    store, embedder, version = Store(), Embedder(), ['a'*64]
    return RagService(store,embedder,lambda doc_id:version[0]),store,embedder,version


def test_structure_preserves_rows_headings_and_table_boundaries():
    blocks = [block('甲'),block('乙',location='paragraph:2'),block('章节',type='heading'),
              block('表一行',type='table_row',group='table:1'),
              block('表一第二行',type='table_row',group='table:1',location='row:2'),
              block('表二行',type='table_row',group='table:2')]
    chunks = structure_chunks(blocks)
    assert [x['text'] for x in chunks] == ['甲\n乙','章节','表一行\n表一第二行','表二行']
    assert 'paragraph:2' in chunks[0]['location']
    assert all(x['kind'] == 'document' for x in chunks)


def test_single_large_paragraph_remains_whole_but_oversize_is_rejected():
    assert structure_chunks([block('甲'*8000)])[0]['text'] == '甲'*8000
    with pytest.raises(RagError):
        structure_chunks([block('甲'*8001)])


@pytest.mark.parametrize('blocks', [[],[None],[block(type=[])],[block(text='')],[block(extra=1)],
    [block(source='.env')],[block(source='dir/.env.local')],[block(text='API_KEY=fixture-private')],
    [block(text='sk-fixturesecret123456')],[block(text='x\x00y')],[block(group='x'*65)]])
def test_invalid_sensitive_and_oversize_blocks_fail_closed(blocks):
    with pytest.raises(RagError):
        structure_chunks(blocks)


def test_capacity_rejects_instead_of_silently_omitting_blocks():
    with pytest.raises(RagError):
        structure_chunks([block('heading',type='heading') for _ in range(501)])
    with pytest.raises(RagError):
        structure_chunks([block() for _ in range(4001)])


def test_location_budget_and_text_budget_preserve_every_input():
    blocks = [block(str(i),location='p'+str(i)+'x'*120) for i in range(100)]
    chunks = structure_chunks(blocks)
    assert '\n'.join(x['text'] for x in chunks) == '\n'.join(str(i) for i in range(100))
    assert all(len(x['location']) <= 600 for x in chunks)


@pytest.mark.parametrize('confirmed',[False,None,1,'true',{},[]])
def test_no_embedding_or_write_without_literal_human_approval(confirmed):
    service,store,embedder,version = setup()
    plan = service.plan_index('document',[block()],'account')
    assert embedder.calls == [] and store.calls == []
    assert service.confirm_index(plan['plan_id'],'account',confirmed)['wrote_nothing'] is True
    assert embedder.calls == [] and store.calls == []
    assert service.confirm_index(plan['plan_id'],'account',True)['status'] == 'rejected'


def test_confirmed_index_uses_exact_planned_immutable_text_and_owned_session():
    service,store,embedder,version = setup()
    blocks = [block('original')]
    plan = service.plan_index('document',blocks,'account')
    blocks[0]['text'] = 'changed'
    assert service.confirm_index(plan['plan_id'],'another',True)['status'] == 'rejected'
    result = service.confirm_index(plan['plan_id'],'account',True)
    assert result['applied'] is True
    assert embedder.calls == [['original']]
    assert store.calls[0][0:3] == ('plan_replace','document','a'*64)
    assert store.calls[-1] == ('confirm_replace','inner','account',True)


@pytest.mark.parametrize('change',['source','database','model'])
def test_changed_source_database_or_model_rejects_before_embedding(change):
    service,store,embedder,version = setup()
    plan = service.plan_index('document',[block()],'account')
    if change == 'source': version[0] = 'b'*64
    if change == 'database': store.metadata['revision'] = 2
    if change == 'model': embedder.identity = {**embedder.identity,'model':'changed-model'}
    assert service.confirm_index(plan['plan_id'],'account',True)['wrote_nothing']
    assert store.calls == [] and embedder.calls == []


def test_source_change_during_embedding_rejects_before_store_plan():
    service,store,embedder,version = setup()
    plan = service.plan_index('document',[block()],'account')
    embedder.on_embed = lambda: version.__setitem__(0,'b'*64)
    assert service.confirm_index(plan['plan_id'],'account',True)['wrote_nothing']
    assert len(embedder.calls) == 1 and store.calls == []


@pytest.mark.parametrize('vectors',[[],[[0,0]],[[float('nan'),1]],[[float('inf'),1]],[[True,0]],[[1]],[[1e7,0]]])
def test_malformed_vectors_never_reach_store_plan(vectors):
    service,store,embedder,version = setup()
    plan = service.plan_index('document',[block()],'account')
    embedder.result = vectors
    assert service.confirm_index(plan['plan_id'],'account',True)['wrote_nothing']
    assert store.calls == []


def test_expiry_and_logout_retire_plans(monkeypatch):
    service,store,embedder,version = setup()
    monkeypatch.setattr('office_agent.rag.time.monotonic',lambda:0)
    plan = service.plan_index('document',[block()],'account')
    monkeypatch.setattr('office_agent.rag.time.monotonic',lambda:301)
    assert service.confirm_index(plan['plan_id'],'account',True)['wrote_nothing']
    plan = service.plan_index('document',[block()],'account')
    service.invalidate()
    assert service.confirm_index(plan['plan_id'],'account',True)['wrote_nothing']
    assert embedder.calls == []


@pytest.mark.parametrize('failure',[True,False])
def test_uncertain_write_never_becomes_success_and_fuse_survives_invalidation(failure):
    service,store,embedder,version = setup()
    plan = service.plan_index('document',[block()],'account')
    store.failure = failure
    store.result = {'status':'uncertain','applied':None,'wrote_nothing':False}
    result = service.confirm_index(plan['plan_id'],'account',True)
    assert result['applied'] is None
    assert 'fixture-private-error' not in str(result)
    service.invalidate()
    with pytest.raises(RagError,match='rag_uncertain'):
        service.plan_index('document',[block()],'account')


def test_query_returns_bounded_excerpt_copy_and_requires_current_source():
    service,store,embedder,version = setup()
    results = service.query('换一种措辞')
    assert len(results[0]['excerpt']) == 200
    assert len(results[0]['text']) == 500
    assert 'excerpt' not in store.hits[0]
    assert store.calls[-1][2] == {'keyword':'换一种措辞','top_k':5,'kind':'document'}
    version[0] = 'b'*64
    with pytest.raises(RagError,match='rag_query_failed'):
        service.query('换一种措辞')


def test_pending_capacity_is_bounded_without_extra_provider_calls():
    service,store,embedder,version = setup()
    for i in range(8):
        service.plan_index(f'doc{i}',[block()],'account')
    with pytest.raises(RagError,match='rag_capacity'):
        service.plan_index('doc9',[block()],'account')
    assert embedder.calls == []
