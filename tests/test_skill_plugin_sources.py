import base64
import hashlib
from io import BytesIO
import json

from PIL import Image
import pytest

from office_agent.public_skill_package import unpack
from office_agent.skills_backend import SkillsBackend
from office_agent.skills_store import SkillsError
from office_agent.webui.public_skill_resolver import resolve
from office_agent.webui.skill_discovery import SkillDiscovery
from tests.test_skill_library import progressive_source


COMMIT = 'a'*40
MIT = b'MIT License\nPermission is hereby granted. Synthetic LiX test fixture.'


def blob(data):
    sha = hashlib.sha1(b'blob '+str(len(data)).encode()+b'\x00'+data).hexdigest()
    return sha, {'sha':sha,'encoding':'base64','content':base64.b64encode(data).decode()}


def plugin_reader(*, dependencies=False, fail_second=False):
    requests, blobs, trees = [], {}, {}
    names = ['reports','reports-extra'] if fail_second else ['reports']
    catalog = {'plugins':[{'name':name,'category':'Productivity','source':{'source':'local','path':'./plugins/'+name},
                           'policy':{'installation':'AVAILABLE','authentication':'ON_INSTALL','products':['CODEX']}} for name in names]}
    catalog['plugins'].append({'name':'external-tool','category':'Security','source':{'source':'url','url':'https://example.invalid/untrusted.git'}})
    raw = json.dumps(catalog).encode()
    sha, market = blob(raw)
    market.update(type='file',path='.agents/plugins/marketplace.json')
    trees[COMMIT] = [{'path':'plugins','type':'tree','mode':'040000','sha':'1'*40}]
    trees['1'*40] = [{'path':name,'type':'tree','mode':'040000','sha':str(index+2)*40} for index,name in enumerate(names)]
    for index,name in enumerate(names):
        manifest = {'name':name,'license':'MIT','skills':'./skills/'}
        if dependencies:
            manifest['mcpServers'] = './.mcp.json'
        source = {'.codex-plugin/plugin.json':json.dumps(manifest).encode(), 'LICENSE':MIT,
                  'skills/report-writing/SKILL.md':b'---\nname: report-writing\ndescription: report writing\n---\nPreserve report numbers.',
                  'unlisted/hidden/SKILL.md':b'---\nname: hidden\ndescription: hidden\n---\nDo not discover.'}
        entries = []
        for path,data in source.items():
            sha,payload = blob(data)
            blobs[sha] = payload
            entries.append({'path':path,'type':'blob','mode':'100644','sha':sha,'size':len(data)})
        trees[str(index+2)*40] = entries
    def read(path,deadline,**options):
        requests.append(path)
        if path.endswith('/commits/HEAD'):
            return COMMIT.encode()
        if '/contents/' in path:
            return json.dumps(market).encode()
        if '/git/trees/' in path:
            identity = path.rsplit('/',1)[-1].split('?')[0]
            if fail_second and identity=='3'*40:
                raise ValueError('source_rate_limited')
            return json.dumps({'tree':trees[identity],'truncated':False}).encode()
        return json.dumps(blobs[path.rsplit('/',1)[-1]]).encode()
    return read,requests


def test_plugin_search_is_scoped_fixed_and_candidate_install_has_no_network(tmp_path):
    read,calls = plugin_reader()
    service = SkillDiscovery()
    result = service.search_plugins('reports',reader=read,commit=COMMIT)
    assert result['ok'] and len(result['items'])==1 and result['matched_plugins']==1
    item = result['items'][0]
    assert item['path']=='plugins/reports/skills/report-writing' and item['commit']==COMMIT
    assert item['plugin_context']['marketplace_policy']['authentication']=='ON_INSTALL'
    checked = service.resolve(item['candidate_id'],reader=read)
    assert checked['ok'] and checked['compatible'] and checked['origin']['plugin_context']['requirements']==[]
    assert not any(path.endswith(COMMIT+'?recursive=1') for path in calls)
    before = len(calls)
    package = service.installation(checked['resolution_id'])
    backend = SkillsBackend(tmp_path)
    plan = backend.plan_public(package,'session')
    assert not list((tmp_path/'skills').glob('*/SKILL.md'))
    assert backend.confirm_change(plan['plan_id'],'session',True)['verified']
    assert backend.read(package['slug'],'user')['origin']['commit']==COMMIT and len(calls)==before


def test_plugin_requirements_and_undeclared_skill_do_not_grant_capabilities():
    read,_ = plugin_reader(dependencies=True)
    checked = resolve('openai/plugins','report-writing',fixed=(COMMIT,'plugins/reports/skills/report-writing'),reader=read)
    assert not checked['compatible'] and checked['_content'] is None
    assert any('mcpServers' in reason for reason in checked['blockers'])
    with pytest.raises(ValueError,match='plugin_skill_not_declared'):
        resolve('openai/plugins','hidden',fixed=(COMMIT,'plugins/reports/unlisted/hidden'),reader=read)


def test_partial_plugin_search_candidates_resolve_and_failures_are_not_empty_success():
    read,_ = plugin_reader(fail_second=True)
    service = SkillDiscovery(fetcher=lambda query:{'skills':[]})
    result = service.search_plugins('reports',reader=read,commit=COMMIT)
    assert not result['ok'] and result['reason']=='plugin_partial' and len(result['items'])==1
    assert result['errors']==[{'plugin':'reports-extra','reason':'source_rate_limited'}]
    assert service.resolve(result['items'][0]['candidate_id'],reader=read)['ok']
    assert service.search('reports')['items']==[]
    def unavailable(*args,**kwargs):
        raise ValueError('source_unavailable')
    cached = service.search_plugins('reports',reader=unavailable,commit=COMMIT)
    assert not cached['ok'] and cached['cached'] and cached['items']==result['items']
    empty = service.search_plugins('no-match',reader=read,commit=COMMIT)
    assert empty['ok'] and empty['items']==[] and not empty['cached']
    external = service.search_plugins('external-tool',reader=read,commit=COMMIT)
    assert not external['ok'] and external['items']==[] and external['errors']==[{'plugin':'external-tool','reason':'plugin_source_unsupported'}]


def raster(color='orange'):
    buffer = BytesIO()
    Image.new('RGB',(64,32),color).save(buffer,format='PNG')
    return buffer.getvalue()


def presentation_source(color='orange'):
    return {'SKILL.md':b'---\nname: report-writing\ndescription: report writing\nlicense: MIT\n---\nPreserve report numbers.',
            'agents/openai.yaml':b'interface:\n  display_name: "Report helper"\n  short_description: "Report review"\n  icon_small: "./assets/icon.png"\n  default_prompt: "DISPLAY_ONLY_NOT_AN_INSTRUCTION"',
            'assets/icon.png':raster(color)}


def package_reader(files,commit=COMMIT):
    contents = {'report-writing/'+path:data for path,data in files.items()}
    contents['LICENSE'] = MIT
    entries,blobs = [],{}
    for path,data in contents.items():
        sha,payload = blob(data)
        entries.append({'path':path,'type':'blob','mode':'100644','sha':sha,'size':len(data)})
        blobs[sha] = payload
    def read(path,deadline,**options):
        if path.endswith('/commits/HEAD'):
            return commit.encode()
        if '/git/trees/' in path:
            return json.dumps({'tree':entries,'truncated':False}).encode()
        return json.dumps(blobs[path.rsplit('/',1)[-1]]).encode()
    return read


def presentation_package(files=None,commit=COMMIT):
    return resolve('lix-tests/workflows','report-writing',fixed=(commit,'report-writing'),
                   reader=package_reader(files or presentation_source(),commit))


def test_display_package_is_preserved_on_reopen_but_never_added_to_model_prompt(tmp_path):
    source = presentation_source()
    source['SKILL.md'] += b' Use [checklist](references/checklist.md).'
    source['references/checklist.md'] = b'Keep report conditions.'
    package = presentation_package(source)
    assert package['compatible'] and 'Package/4' in package['_content']
    backend = SkillsBackend(tmp_path)
    plan = backend.plan_public(package,'session')
    assert backend.confirm_change(plan['plan_id'],'session',False)['wrote_nothing']
    plan = backend.plan_public(package,'session')
    assert backend.confirm_change(plan['plan_id'],'session',True)['verified']
    read = SkillsBackend(tmp_path).read(package['slug'],'user')
    assert read['origin']['presentation']['icon']==package['icon']
    assert read['origin']['presentation']['display']['display_name']=='Report helper'
    assert len(read['origin']['presentation']['files'])==2 and not read['origin']['local_modified']
    prompt = backend.prompt_load(package['slug'],'user')
    assert package['body'] in prompt and 'DISPLAY_ONLY_NOT_AN_INSTRUCTION' not in prompt
    assert 'Keep report conditions.' in prompt
    assert 'base64' not in prompt and 'agents/openai.yaml' not in prompt
    record_body = package['_content'].split('\n---\n',1)[1]
    record,body = record_body.split('\nLiX-Instructions/1\n',1)
    payload = json.loads(record.split('\n',1)[1])
    payload['presentation_files']['assets/icon.png']['sha256']='0'*64
    with pytest.raises(SkillsError):
        unpack('LiX-Public-Skill-Package/4\n'+json.dumps(payload)+'\nLiX-Instructions/1\n'+body)


def test_display_update_edit_and_rollback_preserve_exact_resources(tmp_path):
    backend = SkillsBackend(tmp_path)
    original = presentation_package()
    install = backend.plan_public(original,'session')
    assert backend.confirm_change(install['plan_id'],'session',True)['verified']
    old = backend.read(original['slug'],'user')
    snapshot = backend.public_snapshot(original['slug'],'user')
    service = SkillDiscovery()
    changed = service.resolve_update(snapshot,reader=package_reader(presentation_source('green'),'b'*40))
    assert changed['ok'] and changed['compatible']
    assert changed['update']['files']==[{'path':'assets/icon.png','status':'修改','before':original['manifest']['assets/icon.png'],'after':changed['manifest']['assets/icon.png']}]
    update = backend.plan_public_update(snapshot,service.installation(changed['resolution_id']),'session')
    saved = backend.confirm_change(update['plan_id'],'session',True)
    assert saved['verified'] and backend.read(original['slug'],'user')['origin']['presentation']['icon']!=old['origin']['presentation']['icon']
    undo = backend.plan_rollback(saved['save_id'],'session')
    assert backend.confirm_rollback(undo['rollback_id'],'session',True)['verified']
    assert backend.read(original['slug'],'user')['sha256']==old['sha256']
    edit = backend.plan_change('edit',{'name':original['slug'],'source':'user','content':'---\nname: report-writing\ndescription: report writing\n---\nLocal report guidance.'},'session')
    saved = backend.confirm_change(edit['plan_id'],'session',True)
    reread = backend.read(original['slug'],'user')
    assert saved['verified'] and reread['origin']['local_modified']
    assert reread['origin']['presentation']==old['origin']['presentation']
    assert reread['origin']['format'].endswith('v4')


def test_display_files_fail_closed_for_unsupported_or_corrupt_resources():
    for changes in [
        {'assets/icon.png':b'<svg onload="alert(1)"></svg>'},
        {'agents/openai.yaml':b'interface:\n  icon_small: "../../secret.png"'},
        {'agents/openai.yaml':b'interface:\n  display_name: "Reports"\npolicy:\n  tools: "all"'},
        {'assets/unowned.png':raster()},
        {'scripts/tool.py':b'print(1)'},
        {'agents/openai.yaml':b'interface:\n  icon_small: "./assets/wrong.jpg"','assets/wrong.jpg':raster()},
    ]:
        package = presentation_package({**presentation_source(),**changes})
        assert not package['compatible'] and package['_content'] is None and package['blockers']


def test_display_package_retains_explicit_progressive_reference_loading(tmp_path):
    source = {**presentation_source(),**{name:text.encode() for name,text in progressive_source().items()}}
    service = SkillDiscovery()
    checked = service.resolve(github_url='https://github.com/lix-tests/workflows/tree/'+COMMIT+'/report-writing',reader=package_reader(source))
    assert checked['ok'] and not checked['compatible'] and checked['progressive_available']
    chosen = service.progressive(checked['resolution_id'])
    package = service.installation(chosen['resolution_id'])
    assert chosen['compatible'] and 'Package/4' in package['_content']
    backend = SkillsBackend(tmp_path)
    plan = backend.plan_public(package,'session')
    assert backend.confirm_change(plan['plan_id'],'session',True)['verified']
    read = backend.read(package['slug'],'user')
    assert read['origin']['reference_loading']=='task-matched' and read['origin']['presentation']['icon']==package['icon']
    context = backend.prompt_for_task(package['slug'],'user','newsletter',8192)
    assert 'Newsletter facts to preserve.' in context['body'] and 'Quarterly figures to preserve.' not in context['body']
    assert 'DISPLAY_ONLY_NOT_AN_INSTRUCTION' not in context['body']
