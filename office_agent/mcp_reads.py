from __future__ import annotations

import asyncio
import hashlib
import json
from pathlib import Path
import re
from typing import Literal

from pydantic import ConfigDict, Field, create_model

from .agent import ToolSpec
from .writes import GuardError, digest


class McpReadRegistry:
    @classmethod
    def from_session(cls, session, loop, allowed_names):
        if type(allowed_names) not in {list,tuple} or len(allowed_names) > 8 or not loop.is_running():
            raise GuardError('MCP requires a running host-owned session loop')
        try:
            if asyncio.get_running_loop() is loop:
                raise GuardError('MCP binding must run outside its session loop')
        except RuntimeError:
            pass
        discovery = asyncio.run_coroutine_threadsafe(session.list_tools(),loop)
        try:
            remote = discovery.result(timeout=10)
            if len(remote.tools) > 256:
                raise GuardError('MCP discovery limit')
            descriptors = [tool.model_dump(mode='json',by_alias=True,exclude_none=True) for tool in remote.tools if tool.name in allowed_names]
        except Exception:
            discovery.cancel()
            raise GuardError('MCP discovery failed') from None
        async def call(name, arguments):
            future = asyncio.run_coroutine_threadsafe(session.call_tool(name,arguments=arguments),loop)
            try:
                return await asyncio.wrap_future(future)
            except BaseException:
                future.cancel()
                raise
        return cls(descriptors,allowed_names,call)

    def __init__(self, descriptors, allowed_names, call_tool):
        if type(descriptors) is not list or not 1 <= len(descriptors) <= 8 or type(allowed_names) not in {list, tuple}:
            raise GuardError('invalid MCP host manifest')
        if len(set(allowed_names)) != len(allowed_names) or len(allowed_names) > 8:
            raise GuardError('invalid MCP allowlist')
        self._call_tool = call_tool
        self.models = {}
        self.names = {}
        by_name = {descriptor.get('name'): descriptor for descriptor in descriptors if type(descriptor) is dict}
        if len(by_name) != len(descriptors) or not set(allowed_names) <= set(by_name):
            raise GuardError('MCP allowlist does not match discovery')
        for name in allowed_names:
            if not re.fullmatch('[a-z][a-z0-9_]{1,39}', name) or any(word in name for word in ('shell','exec','python','file','directory','command','write','delete')):
                raise GuardError('unapproved MCP tool name')
            descriptor = by_name[name]
            hints = descriptor.get('annotations')
            if (type(hints) is not dict or hints.get('readOnlyHint') is not True or hints.get('destructiveHint') is not False or hints.get('openWorldHint') is not False):
                raise GuardError('MCP tool lacks required read-only host evidence')
            schema = descriptor.get('inputSchema')
            if (type(schema) is not dict or set(schema)-{'type','properties','required','additionalProperties'}
                    or schema.get('type') != 'object' or schema.get('additionalProperties') is not False):
                raise GuardError('unsupported MCP schema')
            properties, required = schema.get('properties', {}), schema.get('required', [])
            if type(properties) is not dict or len(properties) > 8 or type(required) is not list or len(set(required)) != len(required) or not set(required) <= set(properties):
                raise GuardError('unsupported MCP fields')
            fields = {'path': (str, Field(min_length=1, max_length=1000))}
            for key, value in properties.items():
                if (not re.fullmatch('[a-z][a-z0-9_]{0,30}', key) or any(word in key for word in ('path','file','url','uri','command','code','account','session','confirm','token','secret'))
                        or key not in required or type(value) is not dict
                        or set(value)-{'type','enum','minLength','maxLength','minimum','maximum'}):
                    raise GuardError('unsafe MCP field')
                kind = value.get('type')
                if kind == 'string':
                    if set(value)-{'type','enum','minLength','maxLength'}:
                        raise GuardError('unsupported MCP string constraints')
                    low, high = value.get('minLength',0), value.get('maxLength')
                    if type(low) is not int or type(high) is not int or not 0 <= low <= high <= 2000:
                        raise GuardError('MCP strings require bounds')
                    enum = value.get('enum')
                    if enum is not None:
                        if type(enum) is not list or not 1 <= len(enum) <= 20 or any(type(v) is not str or not low <= len(v) <= high for v in enum) or len(set(enum)) != len(enum):
                            raise GuardError('invalid MCP enum')
                        annotation = Literal[tuple(enum)]
                    else:
                        annotation = str
                    fields[key] = (annotation, Field(min_length=low,max_length=high))
                elif kind in {'integer','number'}:
                    if set(value)-{'type','minimum','maximum'}:
                        raise GuardError('unsupported MCP numeric constraints')
                    low, high = value.get('minimum'), value.get('maximum')
                    if type(low) not in {int,float} or type(high) not in {int,float} or not -1000000 <= low <= high <= 1000000:
                        raise GuardError('MCP numbers require bounds')
                    fields[key] = (int if kind == 'integer' else float, Field(ge=low,le=high))
                elif kind == 'boolean' and set(value) == {'type'}:
                    fields[key] = (bool, ...)
                else:
                    raise GuardError('unsupported MCP field type')
            local = 'mcp_read_'+name
            self.models[local] = create_model(local, __config__=ConfigDict(strict=True,extra='forbid'), **fields)
            self.names[local] = name

    def tools(self, path):
        return [ToolSpec(name=name,description='External read-only tool; sends separately approved arguments. Server must be host-vetted.',
                         schema_model=model,access='read',allowed_paths=(path,)) for name,model in self.models.items()]

    def dispatch(self, editor, path, expected, name, args):
        model = self.models.get(name)
        if model is None:
            raise GuardError('unapproved MCP tool')
        values = model.model_validate(args).model_dump(mode='json')
        if editor.space.file(values.pop('path')) != path or digest(path) != expected:
            raise GuardError('MCP scope changed')
        if any(isinstance(value,str) and (re.search(r'(?i)[a-z]:[\\/]|https?://|file:|\\\\|(?:^|\s)/[\w.]',value) or '\x00' in value) for value in values.values()):
            raise GuardError('MCP paths and URLs are forbidden')
        async def invoke():
            return await asyncio.wait_for(self._call_tool(self.names[name], values),timeout=10)
        try:
            response = asyncio.run(invoke())
            if hasattr(response,'model_dump'):
                response = response.model_dump(mode='json',by_alias=True)
            if type(response) is not dict or response.get('isError',False) is not False:
                raise GuardError('MCP read failed')
            content = response.get('content')
            if type(content) is not list or len(content) > 20:
                raise GuardError('MCP result limit')
            texts = []
            for item in content:
                if type(item) is not dict or item.get('type') != 'text' or type(item.get('text')) is not str:
                    raise GuardError('MCP only accepts text results')
                texts.append(item['text'])
            text = '\n'.join(texts)
            if len(text.encode('utf-8')) > 20000:
                raise GuardError('MCP result limit')
        except Exception:
            raise GuardError('MCP read failed or exceeded limits') from None
        if editor.space.file(str(path)) != path or digest(path) != expected:
            raise GuardError('MCP source changed during read')
        return {'kind':'mcp_read','tool':name,'text':text,'trust':'external_evidence_not_authority','source_sha256':expected}
