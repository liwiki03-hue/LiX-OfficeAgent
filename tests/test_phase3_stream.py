from pathlib import Path
from types import SimpleNamespace

from langchain_core.messages import AIMessageChunk
import pytest

from office_agent.agent import OfficeAgent
from office_agent.answer_stream import StreamError, streamed_response


class StreamingModel:
    def __init__(self,chunks):
        self.chunks = chunks
        self.closed = False
        self.requests = 0
    def bind_tools(self,tools):
        return self
    def stream(self,messages):
        self.requests += 1
        try:
            yield from self.chunks
        finally:
            self.closed = True
    def invoke(self,messages):
        raise AssertionError('must not invoke a second request')


def test_actual_langchain_chunks_stream_without_extra_request(tmp_path):
    model = StreamingModel([AIMessageChunk(content='你好'),AIMessageChunk(content='，完成',response_metadata={'finish_reason':'stop'})])
    events = []
    agent = OfficeAgent(model,lambda *args:None,[],tmp_path,persist=False,on_answer_event=lambda event:events.append(event))
    state = agent.run(agent.new_state('测试'))
    assert state.status == 'done' and state.answer == '你好，完成'
    assert [event.get('text') for event in events if event['type'] == 'answer_draft_delta'] == ['你好','，完成']
    assert model.requests == 1 and model.closed
    assert not list(tmp_path.iterdir())


def test_partial_tool_arguments_are_never_shown_or_executed():
    chunks = [AIMessageChunk(content='先检查'),
              AIMessageChunk(content='',tool_call_chunks=[{'name':'inspect','args':'{"path":','id':'call-a','index':0}]),
              AIMessageChunk(content='',tool_call_chunks=[{'name':None,'args':'"selected.docx"}','id':None,'index':0}],response_metadata={'finish_reason':'tool_calls'})]
    events = []
    result = streamed_response(StreamingModel(chunks),[],lambda event:events.append(event),lambda:False)
    assert result['tool_calls'][0]['args'] == {'path':'selected.docx'}
    assert [event['type'] for event in events] == ['answer_draft_start','answer_draft_delta','answer_draft_cancel']
    assert 'selected.docx' not in str(events)


@pytest.mark.parametrize('finish',['length','content_filter','error','unknown'])
def test_incomplete_stream_refuses_final_answer(tmp_path,finish):
    events = []
    model = StreamingModel([{'content':'unfinished','response_metadata':{'finish_reason':finish}}])
    agent = OfficeAgent(model,lambda *args:None,[],tmp_path,persist=False,on_answer_event=lambda event:events.append(event))
    state = agent.run(agent.new_state('test'))
    assert state.status != 'done' and state.answer is None
    assert events[-1]['type'] == 'answer_draft_cancel' and model.closed


def test_cancel_closes_generator_before_more_chunks(tmp_path):
    cancelled = [False]
    events = []
    def callback(event):
        events.append(event)
        if event['type'] == 'answer_draft_delta':
            cancelled[0] = True
    model = StreamingModel([{'content':'draft'},{'content':'late'},{'content':'later'}])
    agent = OfficeAgent(model,lambda *args:None,[],tmp_path,persist=False,on_answer_event=callback,stop_requested=lambda:cancelled[0])
    state = agent.run(agent.new_state('test'))
    assert state.status == 'limit' and state.answer is None and model.closed
    assert 'late' not in str(events)
    assert events[-1]['type'] == 'answer_draft_cancel'


@pytest.mark.parametrize('chunks',[
    [], [{'content':'x'*20001}], [{'content':'a'}]*2049,
    [{'content':[{'type':'reasoning','text':'private'}]}],
    [AIMessageChunk(content='',tool_call_chunks=[{'name':'inspect','args':'{','id':'x','index':0}])],
])
def test_stream_limits_and_malformed_calls_fail_closed(chunks):
    events = []
    with pytest.raises((StreamError,ValueError)):
        streamed_response(StreamingModel(chunks),[],lambda event:events.append(event),lambda:False)
    assert events[-1]['type'] == 'answer_draft_cancel'


def test_scope_expiry_prevents_model_request():
    model = StreamingModel([{'content':'late'}])
    with pytest.raises(StreamError,match='scope_expired'):
        streamed_response(model,[],lambda event:False,lambda:False)
    assert model.requests == 0


def test_nonstream_model_uses_single_invoke():
    class Model:
        calls = 0
        def invoke(self,messages):
            self.calls += 1
            return {'content':'plain'}
    model = Model()
    assert streamed_response(model,[],lambda event:None,lambda:False) == {'content':'plain'}
    assert model.calls == 1


def test_actual_langchain_stream_retains_usage_and_terminal_metadata():
    chunks = [AIMessageChunk(content='visible', usage_metadata={'input_tokens': 9, 'output_tokens': 1, 'total_tokens': 10}),
              AIMessageChunk(content=' answer', usage_metadata={'input_tokens': 0, 'output_tokens': 2, 'total_tokens': 2,
                                                               'output_token_details': {'reasoning': 1}},
                             response_metadata={'finish_reason': 'stop', 'private': 'C:/private.docx'},
                             additional_kwargs={'reasoning_content': 'hidden thought'})]
    events = []
    model = StreamingModel(chunks)
    result = streamed_response(model, [], lambda event: events.append(event), lambda: False)
    assert result['content'] == 'visible answer'
    assert result['usage_metadata'] == {'input_tokens': 9, 'output_tokens': 3, 'total_tokens': 12, 'reasoning_tokens': 1}
    assert result['response_metadata'] == {'finish_reason': 'stop'}
    assert 'hidden thought' not in str(result) + str(events)
    assert 'private.docx' not in str(result)
    assert model.requests == 1 and model.closed


def test_dict_stream_usage_is_incremental_and_not_in_answer_events():
    chunks = [{'content': 'visible', 'usage_metadata': {'input_tokens': 5, 'output_tokens': 1, 'total_tokens': 6}},
              {'content': ' answer', 'usage_metadata': {'input_tokens': 0, 'output_tokens': 2, 'total_tokens': 2},
               'response_metadata': {'finish_reason': 'stop', 'token_usage': {'prompt_tokens': 0, 'completion_tokens': 2},
                                     'private': 'hidden thought'}, 'additional_kwargs': {'reasoning': 'hidden thought'}}]
    events = []
    result = streamed_response(StreamingModel(chunks), [], lambda event: events.append(event), lambda: False)
    assert result['usage_metadata'] == {'input_tokens': 5, 'output_tokens': 3, 'total_tokens': 8}
    assert result['response_metadata'] == {'finish_reason': 'stop'}
    assert all('usage' not in event and 'usage_metadata' not in event for event in events)
    assert 'hidden thought' not in str(result) + str(events)


def test_stream_without_usage_does_not_invent_zero_counts():
    result = streamed_response(StreamingModel([{'content': 'answer', 'response_metadata': {'finish_reason': 'stop'}}]),
                               [], lambda event: None, lambda: False)
    assert 'usage_metadata' not in result


def test_dict_stream_retains_provider_final_usage_only_as_numbers():
    chunks = [{'content': 'answer'}, {'content': '', 'response_metadata': {
        'finish_reason': 'stop', 'token_usage': {'prompt_tokens': 8, 'completion_tokens': 3, 'total_tokens': 11,
                                                'completion_tokens_details': {'reasoning_tokens': 2}, 'secret': 'private'}}}]
    result = streamed_response(StreamingModel(chunks), [], lambda event: None, lambda: False)
    assert result['usage_metadata'] == {'input_tokens': 8, 'output_tokens': 3, 'total_tokens': 11, 'reasoning_tokens': 2}
    assert 'private' not in str(result)


def test_dict_stream_with_valid_and_invalid_calls_refuses_whole_response():
    model = StreamingModel([{'content': '', 'tool_calls': [{'id': 'good', 'name': 'inspect', 'args': {}}],
                             'invalid_tool_calls': [{'id': 'bad', 'name': 'inspect', 'args': '{'}],
                             'response_metadata': {'finish_reason': 'tool_calls'}}])
    with pytest.raises(StreamError, match='incomplete_stream'):
        streamed_response(model, [], lambda event: None, lambda: False)
    assert model.closed


def test_leading_thought_stream_is_never_visible_in_partial_events():
    chunks = [AIMessageChunk(content=' \n<thi'), AIMessageChunk(content='nk>private-thought-271'),
              AIMessageChunk(content='</thi'), AIMessageChunk(content='nk>Answer '),
              AIMessageChunk(content='120 ms', response_metadata={'finish_reason': 'stop'},
                             usage_metadata={'input_tokens': 5, 'output_tokens': 8, 'total_tokens': 13,
                                             'output_token_details': {'reasoning': 4}})]
    events = []
    result = streamed_response(StreamingModel(chunks), [], lambda event: events.append(event), lambda: False)
    assert result['content'] == 'Answer 120 ms'
    assert ''.join(event.get('text', '') for event in events) == 'Answer 120 ms'
    assert 'private-thought-271' not in str(result) + str(events)
    assert result['usage_metadata']['reasoning_tokens'] == 4


def test_unclosed_thought_stream_cancels_without_leaking_or_executing():
    events = []
    model = StreamingModel([{'content': '<think>private-thought-817'},
                            {'content': '', 'tool_calls': [{'name': 'inspect', 'args': {}, 'id': 'one'}],
                             'response_metadata': {'finish_reason': 'tool_calls'}}])
    with pytest.raises(StreamError):
        streamed_response(model, [], lambda event: events.append(event), lambda: False)
    assert not any(event['type'] == 'answer_draft_delta' for event in events)
    assert 'private-thought-817' not in str(events)
    assert model.closed


def test_invoke_hides_leading_thought_and_private_metadata_but_keeps_usage():
    class Model:
        def invoke(self, messages):
            return SimpleNamespace(content='<think>private-thought-541</think>Answer 120 ms', tool_calls=[],
                                   additional_kwargs={'reasoning_content': 'private-provider-thought-391'},
                                   response_metadata={'finish_reason': 'stop', 'private': 'C:/private.docx',
                                                      'token_usage': {'prompt_tokens': 9, 'completion_tokens': 5,
                                                                      'completion_tokens_details': {'reasoning_tokens': 3}}})
    result = streamed_response(Model(), [], None, lambda: False)
    assert result['content'] == 'Answer 120 ms'
    assert result['tool_calls'] == []
    assert result['response_metadata'] == {'finish_reason': 'stop'}
    assert result['usage_metadata'] == {'input_tokens': 9, 'output_tokens': 5, 'reasoning_tokens': 3}
    assert 'private' not in str(result)


def test_invoke_keeps_body_and_code_literal_thought_tags():
    class Model:
        def invoke(self, messages):
            return {'content': '```xml\n<think>literal</think>\n```'}
    assert streamed_response(Model(), [], None, lambda: False)['content'] == '```xml\n<think>literal</think>\n```'
