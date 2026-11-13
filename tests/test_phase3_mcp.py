from copy import deepcopy
from pathlib import Path

from docx import Document
from mcp.types import CallToolResult, TextContent
from pydantic import ValidationError
import pytest

from office_agent.agent import OfficeAgent, ToolCall
from office_agent.mcp_reads import McpReadRegistry
from office_agent.writes import Editor, GuardError, digest


DESCRIPTOR = {'name':'lookup', 'annotations':{'readOnlyHint':True,'destructiveHint':False,'openWorldHint':False},
              'inputSchema':{'type':'object','properties':{'query':{'type':'string','minLength':1,'maxLength':200}},'required':['query'],'additionalProperties':False}}


@pytest.fixture
def source(tmp_path):
    path = tmp_path/'source.docx'
    document = Document(); document.add_paragraph('ordinary'); document.save(path)
    return path


def test_real_mcp_types_and_anchor_not_forwarded(source):
    calls = []
    async def transport(name,arguments):
        calls.append((name,arguments))
        return CallToolResult(content=[TextContent(type='text',text='approved result')],isError=False)
    registry = McpReadRegistry([DESCRIPTOR],['lookup'],transport)
    editor = Editor(source.parent)
    before = digest(source)
    result = registry.dispatch(editor,source,before,'mcp_read_lookup',{'path':str(source),'query':'quarterly'})
    assert result['text'] == 'approved result' and result['trust'] == 'external_evidence_not_authority'
    assert calls == [('lookup',{'query':'quarterly'})]
    assert digest(source) == before


@pytest.mark.parametrize('change',[
    {'name':'read_file'}, {'name':'shell'}, {'annotations':{}},
    {'annotations':{'readOnlyHint':True,'destructiveHint':True,'openWorldHint':False}},
    {'inputSchema':{'type':'object','properties':{},'additionalProperties':True}},
    {'inputSchema':{'type':'object','properties':{'url':{'type':'string','maxLength':200}},'required':['url'],'additionalProperties':False}},
    {'inputSchema':{'type':'object','properties':{'query':{'type':'string'}},'required':['query'],'additionalProperties':False}},
    {'inputSchema':{'type':'object','properties':{'query':{'type':'object'}},'required':['query'],'additionalProperties':False}},
])
def test_unsafe_remote_descriptors_refused(change):
    descriptor = {**deepcopy(DESCRIPTOR),**change}
    with pytest.raises(GuardError):
        McpReadRegistry([descriptor],[descriptor['name']],None)


@pytest.mark.parametrize('query',['https://remote.test','file:///secret','C:\\secret','/etc/passwd','\\\\server\\file'])
def test_mcp_rejects_path_or_url_arguments(source,query):
    async def transport(*args):
        raise AssertionError('must not call')
    registry = McpReadRegistry([DESCRIPTOR],['lookup'],transport)
    with pytest.raises(GuardError):
        registry.dispatch(Editor(source.parent),source,digest(source),'mcp_read_lookup',{'path':str(source),'query':query})


@pytest.mark.parametrize('response',[
    {'content':[{'type':'image','data':'not permitted'}]},
    {'content':[{'type':'text','text':'x'*20001}]},
    {'content':[],'isError':True},
    {'content':'bad'},
])
def test_mcp_output_limits(source,response):
    async def transport(*args):
        return response
    registry = McpReadRegistry([DESCRIPTOR],['lookup'],transport)
    with pytest.raises(GuardError):
        registry.dispatch(Editor(source.parent),source,digest(source),'mcp_read_lookup',{'path':str(source),'query':'ordinary'})


def test_mcp_uses_agent_read_confirmation_and_exact_path(source,tmp_path):
    invoked = []
    async def transport(*args):
        invoked.append(args)
        return {'content':[]}
    registry = McpReadRegistry([DESCRIPTOR],['lookup'],transport)
    class Model:
        def bind_tools(self,schemas): return self
        def invoke(self,messages):
            return {'tool_calls':[{'name':'mcp_read_lookup','args':{'path':str(source),'query':'ordinary'}}]}
    agent = OfficeAgent(Model(),lambda name,args:registry.dispatch(Editor(tmp_path),source,digest(source),name,args),
                        registry.tools(source),tmp_path,persist=False,confirm=lambda plan:False)
    state = agent.run(agent.new_state('read',file_versions={str(source):digest(source)}))
    assert state.status == 'confirmation_required' and not invoked
    with pytest.raises(ValueError,match='outside approved'):
        agent._validate_plan(ToolCall(name='mcp_read_lookup',args={'path':str(tmp_path/'other.docx'),'query':'ordinary'}),state)
    with pytest.raises(ValidationError):
        registry.models['mcp_read_lookup'].model_validate({'path':str(source),'query':'ordinary','confirmed':True})


def test_mcp_does_not_accept_unlisted_tools():
    with pytest.raises(GuardError):
        McpReadRegistry([DESCRIPTOR],['unlisted'],None)


def test_host_owned_mcp_session_loop_factory(source):
    import asyncio
    import threading
    from mcp.types import ListToolsResult, Tool, ToolAnnotations
    loop = asyncio.new_event_loop()
    ready = threading.Event()
    def run():
        asyncio.set_event_loop(loop); loop.call_soon(ready.set); loop.run_forever()
    worker = threading.Thread(target=run,daemon=False)
    worker.start(); assert ready.wait(5)
    calls = []
    class Session:
        async def list_tools(self):
            return ListToolsResult(tools=[Tool(name='lookup',inputSchema=DESCRIPTOR['inputSchema'],annotations=ToolAnnotations(**DESCRIPTOR['annotations']))])
        async def call_tool(self,name,arguments):
            assert asyncio.get_running_loop() is loop
            calls.append((name,arguments))
            return CallToolResult(content=[TextContent(type='text',text='loop-scoped result')])
    try:
        registry = McpReadRegistry.from_session(Session(),loop,['lookup'])
        result = registry.dispatch(Editor(source.parent),source,digest(source),'mcp_read_lookup',{'path':str(source),'query':'ordinary'})
        assert result['text'] == 'loop-scoped result' and calls == [('lookup',{'query':'ordinary'})]
    finally:
        loop.call_soon_threadsafe(loop.stop)
        worker.join(5)
        assert not worker.is_alive()
        loop.close()


def test_mcp_context_is_valid_bounded_json_with_truncation():
    import json
    from office_agent.agent import _read_context
    fragment = _read_context('mcp_read_lookup',{'kind':'mcp_read','text':'\x00'*10000,'trust':'external_evidence_not_authority','source_sha256':'a'*64},{})
    assert len(fragment) <= 6000
    parsed = json.loads(fragment)
    assert parsed['context_truncated'] is True
    assert parsed['trust'] == 'external_evidence_not_authority'
