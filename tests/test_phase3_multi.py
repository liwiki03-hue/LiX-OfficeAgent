from pathlib import Path

from docx import Document
from openpyxl import Workbook, load_workbook
import pytest

from office_agent.multi_files import MultiFileScope, MultiFileSelection
from office_agent.writes import Editor, GuardError, digest


@pytest.fixture
def sources(tmp_path):
    word = tmp_path/'one.docx'
    doc = Document(); doc.add_paragraph('Original'); doc.save(word)
    excel = tmp_path/'two.xlsx'
    book = Workbook(); book.active['A1'] = 1; book.save(excel)
    third = tmp_path/'third.docx'; third.write_bytes(word.read_bytes())
    return word, excel, third


def test_multi_summary_is_bounded_explicit_and_read_only(sources,tmp_path):
    word,excel,_ = sources
    versions = {path:digest(path) for path in sources}
    scope = MultiFileScope(Editor(tmp_path),word,[word.name,excel.name])
    result = scope.summary()
    assert len(result['documents']) == 2
    assert {item['source'] for item in result['documents']} == {word.name,excel.name}
    assert all(len(item['excerpts']) <= 5 for item in result['documents'])
    assert {path:digest(path) for path in sources} == versions
    assert not (tmp_path/'.office_agent').exists()


@pytest.mark.parametrize('names',[['one.docx'],['one.docx']*2,['one.docx','../two.xlsx'],['one.docx','absent.docx'],['two.xlsx','third.docx'],['one.docx']*6])
def test_multi_scope_invalid_input_fails_closed(sources,tmp_path,names):
    with pytest.raises((GuardError,ValueError)):
        MultiFileScope(Editor(tmp_path),sources[0],names)
    assert not (tmp_path/'.office_agent').exists()


@pytest.mark.parametrize('confirmed',[False,1,'true'])
def test_multi_initial_read_needs_strict_confirmation(sources,tmp_path,confirmed):
    selection = MultiFileSelection()
    editor = Editor(tmp_path)
    plan = selection.plan(editor,sources[0],[p.name for p in sources[:2]])
    with pytest.raises(GuardError):
        selection.confirm(editor,sources[0],plan['plan_id'],confirmed)
    assert selection.scope is None


def test_each_write_is_its_own_plan_confirmation_and_rollback(sources,tmp_path):
    word,excel,other = sources
    editor = Editor(tmp_path,require_confirmation=True)
    scope = MultiFileScope(editor,word,[word.name,excel.name])
    dispatch = scope.dispatcher()
    word_plan = dispatch('plan_word',{'path':str(word),'paragraph_index':0,'old':'Original','new':'Updated'})
    excel_plan = dispatch('plan_excel',{'path':str(excel),'sheet':'Sheet','cell':'A1','value':5})
    for result in (word_plan,excel_plan):
        with pytest.raises(GuardError):
            editor.execute(result['plan_id'])
    assert not (tmp_path/'.office_agent').exists()
    first = editor.execute(word_plan['plan_id'],confirmed=True)
    assert excel_plan['plan_id'] in editor.plans
    assert load_workbook(excel).active['A1'].value == 1
    second = editor.execute(excel_plan['plan_id'],confirmed=True)
    assert Document(word).paragraphs[0].text == 'Original'
    assert digest(word) == scope.versions[str(word)] and digest(excel) == scope.versions[str(excel)]
    assert first['source'] != second['source']
    assert editor.rollback(first['id'],confirmed=True)['status'] == 'rolled_back'
    assert Path(second['output']).exists()
    assert editor.rollback(second['id'],confirmed=True)['status'] == 'rolled_back'
    with pytest.raises(GuardError):
        dispatch('inspect',{'path':str(other)})


def test_any_input_change_invalidates_whole_multi_scope(sources,tmp_path):
    scope = MultiFileScope(Editor(tmp_path),sources[0],[p.name for p in sources[:2]])
    book = load_workbook(sources[1]); book.active['A1'] = 88; book.save(sources[1])
    with pytest.raises(GuardError,match='scope changed'):
        scope.dispatcher()('inspect',{'path':str(sources[0])})


@pytest.mark.parametrize('name',['batch_documents','search_directory_documents','image_gen','plan_docx_merge','plan_docx_split'])
def test_directory_and_multi_source_writes_not_exposed_in_scope(sources,tmp_path,name):
    scope = MultiFileScope(Editor(tmp_path),sources[0],[p.name for p in sources[:2]])
    with pytest.raises(GuardError):
        scope.dispatcher()(name,{'path':str(sources[0])})
