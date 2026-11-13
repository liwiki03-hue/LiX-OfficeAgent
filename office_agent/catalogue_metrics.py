from __future__ import annotations

import json
from pathlib import Path

from .agent import OfficeAgent
from .app import build_tools, select_tools, split_read_write


def catalogue_metrics(path):
    tools = build_tools((path,),ocr_enabled=True)
    class Model:
        def bind_tools(self,*args):
            raise AssertionError('metrics must never call a model')
    full = OfficeAgent(Model(),lambda *args:None,tools,path.parent,persist=False,context_tokens=32768,reply_tokens=768)
    baseline = full.fixed_overhead()
    result = {'estimator':'ceil(UTF-8 bytes / 3); approximate, not provider tokens',
              'baseline_tool_count':len(tools),'baseline_fixed_tokens':baseline,
              'schema_tokens':{tool.name:(len(json.dumps(tool.model_schema(),ensure_ascii=False).encode('utf-8'))+2)//3 for tool in tools},
              'initial_packs':{}}
    for kind in ('word','spreadsheet'):
        read,write = split_read_write(select_tools(tools,kind))
        agent = OfficeAgent(Model(),lambda *args:None,read,path.parent,persist=False,context_tokens=32768,reply_tokens=768,
                            write_tools=write,write_gate='request_write_tools',capability_tools=tools)
        cost = agent.fixed_overhead()
        result['initial_packs'][kind] = {'fixed_tokens':cost,'cost_ratio':cost/baseline,
            'tool_count':len(agent.tools),'count_ratio':len(agent.tools)/len(tools),'visible_names':sorted(agent.tools)}
    result['below_half'] = all(pack['cost_ratio'] < .5 and pack['count_ratio'] < .5 for pack in result['initial_packs'].values())
    return result
