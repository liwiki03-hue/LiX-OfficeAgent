import base64
import hashlib
import json
import re
import time

import httpx2


REPOSITORY = 'openai/plugins'
SHA = re.compile(r'[0-9a-f]{40}\Z')
NAME = re.compile(r'[a-z0-9]+(?:-[a-z0-9]+)*\Z')
LICENSES = {'LICENSE','LICENSE.txt','LICENSE.md','NOTICE','NOTICE.txt','NOTICE.md'}


def decoded_blob(payload, limit):
    if not isinstance(payload,dict) or payload.get('encoding')!='base64' or not SHA.fullmatch(payload.get('sha','')):
        raise ValueError('source_format')
    data = base64.b64decode(payload['content'].replace('\n',''),validate=True)
    if len(data)>limit or hashlib.sha1(b'blob '+str(len(data)).encode()+b'\x00'+data).hexdigest()!=payload['sha']:
        raise ValueError('source_integrity')
    return data


def contents(reader, path, commit, deadline, limit=65536):
    payload = json.loads(reader('/repos/'+REPOSITORY+'/contents/'+path+'?ref='+commit,deadline,limit=limit*2))
    if not isinstance(payload,dict) or payload.get('type')!='file' or payload.get('path')!=path:
        raise ValueError('source_format')
    return decoded_blob(payload,limit)


def tree(reader, identity, deadline, recursive=False):
    payload = json.loads(reader('/repos/'+REPOSITORY+'/git/trees/'+identity+('?recursive=1' if recursive else ''),deadline))
    if not isinstance(payload,dict) or payload.get('truncated') is not False or not isinstance(payload.get('tree'),list) or len(payload['tree'])>5000:
        raise ValueError('source_tree_budget')
    seen = set()
    for item in payload['tree']:
        if not isinstance(item,dict):
            raise ValueError('source_format')
        path = item.get('path','')
        if not isinstance(path,str) or len(path)>512 or len(path.split('/'))>12 or any(not re.fullmatch(r'[A-Za-z0-9_.-]+',part) or part in ('.','..') for part in path.split('/')) or path.lower() in seen:
            raise ValueError('unsupported_source_path')
        if not recursive and '/' in path:
            raise ValueError('source_format')
        seen.add(path.lower())
    return payload['tree']


def declared_paths(value):
    paths = [value] if isinstance(value,str) else value
    if not isinstance(paths,list) or not 1<=len(paths)<=4:
        raise ValueError('plugin_manifest_unsupported')
    result = []
    for path in paths:
        if not isinstance(path,str) or not path.startswith('./skills') or not re.fullmatch(r'\./skills(?:/[A-Za-z0-9_.-]+)*/?',path) or any(part in ('.','..') for part in path[2:].split('/')):
            raise ValueError('plugin_manifest_unsupported')
        result.append(path[2:].rstrip('/'))
    return result


def scoped_plugin(reader, commit, plugin, deadline):
    if not SHA.fullmatch(commit) or not NAME.fullmatch(plugin):
        raise ValueError('plugin_manifest_unsupported')
    ancestors = []
    rows = tree(reader,commit,deadline)
    prefix = ''
    for segment in ('plugins',plugin):
        ancestors.extend({**item,'path':prefix+item['path']} for item in rows if item['path'] in LICENSES)
        match = next((item for item in rows if item['path']==segment),None)
        if not match or match.get('type')!='tree' or match.get('mode') not in ('040000','40000') or not SHA.fullmatch(match.get('sha','')):
            raise ValueError('plugin_path_missing')
        prefix += segment+'/'
        rows = tree(reader,match['sha'],deadline,recursive=segment==plugin)
    paths = [{**item,'path':prefix+item['path']} for item in rows]
    entry = next((item for item in paths if item['path']==prefix+'.codex-plugin/plugin.json'),None)
    if not entry or entry.get('mode')!='100644' or entry.get('type')!='blob':
        raise ValueError('plugin_manifest_unsupported')
    payload = json.loads(reader('/repos/'+REPOSITORY+'/git/blobs/'+entry['sha'],deadline,limit=65536))
    raw = decoded_blob(payload,32768)
    if hashlib.sha1(b'blob '+str(len(raw)).encode()+b'\x00'+raw).hexdigest()!=entry['sha']:
        raise ValueError('source_integrity')
    manifest = json.loads(raw)
    if not isinstance(manifest,dict) or manifest.get('name')!=plugin:
        raise ValueError('plugin_manifest_unsupported')
    roots = declared_paths(manifest.get('skills'))
    license_name = manifest.get('license')
    if license_name is not None and (not isinstance(license_name,str) or len(license_name)>200):
        raise ValueError('plugin_manifest_unsupported')
    requirements = [key for key in ('apps','mcpServers','hooks','commands') if manifest.get(key)]
    requirements.extend(item['path'][len(prefix):] for item in paths if item['path'] in {prefix+'.mcp.json',prefix+'.app.json',prefix+'hooks.json'})
    context = {'name':plugin,'manifest_sha256':hashlib.sha256(raw).hexdigest(),
               'requirements':sorted(set(requirements)),'declared_skills':roots,
               'declared_license':license_name,
               'scope':'仅独立Skill；不安装插件、连接器、hooks或认证'}
    return {'truncated':False,'tree':ancestors+paths}, context


def search(query, reader, *, commit=None):
    started = time.monotonic()
    deadline = started+35
    commit = commit or reader('/repos/'+REPOSITORY+'/commits/HEAD',deadline,accept='application/vnd.github.sha',limit=100).decode().strip()
    if not SHA.fullmatch(commit):
        raise ValueError('invalid_source_version')
    raw = contents(reader,'.agents/plugins/marketplace.json',commit,deadline,262144)
    catalog = json.loads(raw)
    if not isinstance(catalog,dict) or not isinstance(catalog.get('plugins'),list) or len(catalog['plugins'])>128:
        raise ValueError('plugin_manifest_unsupported')
    matches,unsupported,seen,matched = [],[],set(),0
    for item in catalog['plugins']:
        if not isinstance(item,dict):
            raise ValueError('plugin_manifest_unsupported')
        name = item.get('name','')
        if not isinstance(name,str) or not NAME.fullmatch(name) or name in seen:
            raise ValueError('plugin_manifest_unsupported')
        seen.add(name)
        category = item.get('category','')
        if not isinstance(category,str) or len(category)>100:
            raise ValueError('plugin_manifest_unsupported')
        policy = item.get('policy',{})
        if not isinstance(policy,dict) or len(policy)>8:
            raise ValueError('plugin_manifest_unsupported')
        for key,value in policy.items():
            values = value if isinstance(value,list) else [value]
            if not isinstance(key,str) or len(key)>100 or not 1<=len(values)<=8 or any(not isinstance(entry,str) or len(entry)>100 for entry in values):
                raise ValueError('plugin_manifest_unsupported')
        if query.lower() in (name+' '+category).lower():
            matched += 1
            if item.get('source')=={'source':'local','path':'./plugins/'+name}:
                matches.append(item)
            else:
                unsupported.append({'plugin':name,'reason':'plugin_source_unsupported'})
    items,errors = [],unsupported
    limited = False
    for plugin in sorted(matches,key=lambda item:item['name'])[:2]:
        try:
            scoped,context = scoped_plugin(reader,commit,plugin['name'],deadline)
            context['marketplace_policy'] = plugin.get('policy',{})
            prefix = 'plugins/'+plugin['name']+'/'
            for item in scoped['tree']:
                path = item['path']
                if not path.endswith('/SKILL.md') or not any(path.startswith(prefix+root+'/') for root in context['declared_skills']):
                    continue
                if item.get('type')!='blob' or item.get('mode')!='100644':
                    continue
                if len(items)>=24:
                    limited = True
                    break
                folder = path[:-9]
                items.append({'candidate_id':hashlib.sha256((REPOSITORY+'@'+commit+'/'+folder).encode()).hexdigest()[:32],
                    'name':folder.split('/')[-1],'source':REPOSITORY,'provider':'openai/plugins','skill_id':folder.split('/')[-1],
                    'path':folder,'commit':commit,'url':'https://github.com/'+REPOSITORY+'/blob/'+commit+'/'+path,
                    'description':plugin['name']+' · '+plugin.get('category',''),'compatibility':'not_checked',
                    'plugin_context':context,'license':None})
        except (ValueError,TypeError,KeyError,httpx2.HTTPError) as exc:
            code = 'source_timeout' if isinstance(exc,httpx2.TimeoutException) else str(exc) if isinstance(exc,ValueError) else 'source_unavailable'
            errors.append({'plugin':plugin['name'],'reason':code if code in {'source_tree_budget','plugin_path_missing','plugin_manifest_unsupported','source_integrity','source_rate_limited','source_unavailable','source_timeout','source_budget','source_format','unsupported_source_path'} else 'source_unavailable'})
    return {'ok':not errors,'reason':'plugin_partial' if errors else None,'items':items,'provider':REPOSITORY,
            'query':query,'commit':commit,'manifest_sha256':hashlib.sha256(raw).hexdigest(),
            'matched_plugins':matched,'scanned_plugins':min(2,len(matches)),'limited':len(matches)>2 or limited,
            'errors':errors,'elapsed_ms':round((time.monotonic()-started)*1000),'cached':False,
            'scope':'按插件名称/类别查询，最多2个插件/24项Skill；不是GitHub全文搜索'}
