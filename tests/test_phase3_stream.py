from pathlib import Path

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
