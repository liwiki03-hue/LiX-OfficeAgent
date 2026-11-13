import json

from office_agent.agent import OfficeAgent, _read_context
from office_agent.app import build_dispatcher, build_tools
from office_agent.writes import Editor, digest
from test_read_coverage import make_pptx, word


def test_pptx_model_context_keeps_whole_items_and_exact_continuation():
    items = [{'type': 'text_box', 'item_index': index, 'text': 'x' * 1000,
              'text_offset': 0, 'text_truncated': False} for index in range(20)]
    raw = {'slide_index': 2, 'items': items, 'next_offset': None, 'truncated': False}
    text = _read_context('read_pptx_slide', raw, {'offset': 0})
    view = json.loads(text)
    assert len(text) <= 6000 and view['context_truncated'] is True
    count = len(view['items'])
    assert 0 < count < 20 and view['items'] == items[:count]
    assert view['next_offset'] == count and view['next_text_offset'] == 0


def test_pdf_model_context_has_coordinates_and_progress_by_pdf_character_index():
    runs = [{'char_index': index * 80, 'char_count': 80, 'bbox': [1, 2, 3, 4], 'text': 'z' * 80}
            for index in range(200)]
    raw = {'page_count': 1, 'truncated': False, 'next_page': None, 'next_char_offset': None,
           'pages': [{'page': 0, 'size_points': [200, 200], 'char_offset': 0, 'total_chars': 16000,
                      'text_truncated': False, 'text_layer_empty': False, 'runs': runs, 'text': 'z' * 16000}]}
    text = _read_context('read_pdf', raw, {})
    view = json.loads(text)
    count = len(view['pages'][0]['runs'])
    assert len(text) <= 6000 and 0 < count < 200
    assert view['context_truncated'] and view['next_page'] == 0 and view['next_char_offset'] == count * 80
    assert view['pages'][0]['text_provided_in_runs']
    assert view['pages'][0]['runs'] == runs[:count]


def test_pdf_model_context_requests_smaller_atomic_text_read():
    raw = {'pages': [{'page': 4, 'size_points': [200, 200], 'char_offset': 200, 'total_chars': 20000,
                       'text_truncated': False, 'text_layer_empty': False, 'runs': [], 'text': 'x' * 15000}]}
    view = json.loads(_read_context('read_pdf', raw, {}))
    assert view['pages'] == [] and view['context_truncated']
    assert view['next_page'] == 4 and view['next_char_offset'] == 200 and view['retry_max_chars'] == 2000


def test_directory_model_context_preserves_sources_and_marks_partial_hits():
    hits = [{'source': f'{index}.docx', 'location': 'body:1/paragraph', 'excerpt': 'x' * 1000} for index in range(10)]
    text = _read_context('search_directory_documents', {'hits': hits, 'semantic_available': False}, {})
    view = json.loads(text)
    assert len(text) <= 6000 and 0 < len(view['hits']) < 10 and view['context_truncated']
    assert view['hits'] == hits[:len(view['hits'])] and view['semantic_available'] is False


def test_actual_pptx_dispatch_reaches_model_as_evidence_not_just_a_metadata_prefix(tmp_path):
    source = make_pptx(tmp_path / 'slides.pptx')
    original = digest(source)

    class Model:
        calls = 0

        def bind_tools(self, schemas):
            return self

        def invoke(self, messages):
            self.calls += 1
            if self.calls == 1:
                return {'tool_calls': [{'name': 'read_pptx_slide', 'args': {'path': str(source)}}]}
            evidence = [message['content'] for message in messages if message['role'] == 'tool']
            return {'content': '7319' if evidence and 'Budget 7319' in evidence[-1] else 'missing evidence'}

    tools = [tool for tool in build_tools((source,)) if tool.name == 'read_pptx_slide']
    editor = Editor(tmp_path)
    model = Model()
    agent = OfficeAgent(model, build_dispatcher(editor, source), tools, tmp_path / 'state',
                        persist=False, confirm=lambda _: True, context_tokens=32768)
    state = agent.run(agent.new_state('What budget is stated?'))
    assert state.status == 'done' and state.answer == '7319'
    assert model.calls == 2 and digest(source) == original and not editor.plans
