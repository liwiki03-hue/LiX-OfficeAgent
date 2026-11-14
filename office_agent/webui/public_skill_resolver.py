import base64
import hashlib
import json
import re
import time
from pathlib import PurePosixPath
from urllib.parse import quote, urlsplit

import httpx2

from ..public_skill_package import pack, parse_public, prompt_body
from ..skill_references import MAX_FILES, MAX_SOURCE_BYTES, MAX_REFERENCE_BYTES, expand, safe_name
from ..skill_metadata import presentation
from ..skill_presentation import MAX_ICON_BYTES, SIDECAR, icon_paths, encode_files, validate_files


PART = re.compile(r'[A-Za-z0-9][A-Za-z0-9_.-]{0,99}\Z')
SHA = re.compile(r'[0-9a-f]{40}\Z')
LICENSES = {'LICENSE', 'LICENSE.md', 'LICENSE.txt', 'NOTICE', 'NOTICE.md', 'NOTICE.txt'}


def read_api(path, deadline, *, accept='application/vnd.github+json', limit=1024 * 1024):
    if not path.startswith('/repos/') or time.monotonic() >= deadline:
        raise ValueError('source_timeout')
    from ..network import transport
    with httpx2.Client(timeout=min(8, max(.1, deadline - time.monotonic())), follow_redirects=False, trust_env=False, transport=transport('https://api.github.com')) as client:
        with client.stream('GET', 'https://api.github.com' + path,
                           headers={'Accept': accept, 'User-Agent': 'LiX-PublicSkills/1'}) as response:
            if response.status_code in (403, 429):
                raise ValueError('source_rate_limited')
            if response.status_code != 200:
                raise ValueError('source_unavailable')
            data = bytearray()
            for block in response.iter_bytes():
                data.extend(block)
                if len(data) > limit or time.monotonic() >= deadline:
                    raise ValueError('source_budget')
    return bytes(data)


def github_source(url):
    parsed = urlsplit(url)
    if parsed.scheme != 'https' or parsed.netloc != 'github.com' or parsed.query or parsed.fragment:
        raise ValueError('invalid_github_source')
    parts = parsed.path.strip('/').split('/')
    if len(parts) < 5 or parts[2] not in ('tree', 'blob') or not SHA.fullmatch(parts[3]):
        raise ValueError('fixed_commit_required')
    if not all(PART.fullmatch(part) for part in parts[:2]) or not all(re.fullmatch(r'[A-Za-z0-9_.-]+',part) and part not in ('.','..') for part in parts[4:]):
        raise ValueError('invalid_github_source')
    path = '/'.join(parts[4:])
    if parts[2] == 'blob':
        if not path.endswith('/SKILL.md'):
            raise ValueError('invalid_github_source')
        path = path[:-9]
    return '/'.join(parts[:2]), parts[3], path


def resolve(repository, skill_id, *, fixed=None, folder_path=None, reader=read_api):
    deadline = time.monotonic() + 35
    base = '/repos/' + repository
    commit = fixed[0] if fixed else reader(base + '/commits/HEAD', deadline, accept='application/vnd.github.sha', limit=100).decode().strip()
    if not SHA.fullmatch(commit):
        raise ValueError('invalid_source_version')
    plugin_context = None
    requested_folder = fixed[1] if fixed else folder_path
    if repository=='openai/plugins' and requested_folder:
        from .plugin_sources import scoped_plugin
        match = re.fullmatch(r'plugins/([a-z0-9]+(?:-[a-z0-9]+)*)/(.+)',requested_folder)
        if not match:
            raise ValueError('plugin_path_missing')
        tree,plugin_context = scoped_plugin(reader,commit,match[1],deadline)
        if not any(requested_folder.startswith('plugins/'+match[1]+'/'+path+'/') for path in plugin_context['declared_skills']):
            raise ValueError('plugin_skill_not_declared')
    else:
        tree = json.loads(reader(base + '/git/trees/' + commit + '?recursive=1', deadline))
    if not isinstance(tree, dict) or tree.get('truncated') is not False or not isinstance(tree.get('tree'), list) or len(tree['tree']) > 5000:
        raise ValueError('source_tree_budget')
    entries = {}
    for item in tree['tree']:
        path = item.get('path', '')
        if not isinstance(path, str) or len(path) > 512 or len(path.split('/')) > 12 or any(not re.fullmatch(r'[A-Za-z0-9_.-]+', part) or part in ('.','..') for part in path.split('/')):
            raise ValueError('unsupported_source_path')
        if path.lower() in entries:
            raise ValueError('source_path_collision')
        entries[path.lower()] = item
    matches = [item['path'][:-9] for item in entries.values() if item['path'].endswith('/SKILL.md') and PurePosixPath(item['path']).parent.name == skill_id]
    folder = fixed[1] if fixed else folder_path if folder_path is not None else matches[0] if len(matches) == 1 else None
    if not folder:
        raise ValueError('skill_path_ambiguous')
    if folder.lower() + '/skill.md' not in entries:
        raise ValueError('skill_path_missing')
    package = [item for item in entries.values() if item['path'].startswith(folder + '/') and item.get('type') != 'tree']
    if len(package) > MAX_FILES or sum(item.get('size', 0) for item in package) > MAX_SOURCE_BYTES:
        raise ValueError('package_budget')
    unsafe = [item['path'] for item in package if item.get('mode') != '100644' or item.get('type') != 'blob']
    extras = [item['path'] for item in package if item['path'] != folder + '/SKILL.md' and PurePosixPath(item['path']).name not in LICENSES]
    for item in package:
        safe_name(item['path'][len(folder)+1:])
    def blob(item, limit=65536):
        if item.get('type') != 'blob' or item.get('mode') != '100644' or not SHA.fullmatch(item.get('sha','')):
            raise ValueError('unsupported_file_type')
        data = json.loads(reader(base + '/git/blobs/' + item['sha'], deadline, limit=limit * 2))
        if data.get('encoding') != 'base64':
            raise ValueError('source_format')
        decoded = base64.b64decode(data['content'].replace('\n',''), validate=True)
        if len(decoded) > limit or hashlib.sha1(b'blob ' + str(len(decoded)).encode() + b'\x00' + decoded).hexdigest() != item['sha']:
            raise ValueError('source_integrity')
        return decoded
    original = blob(entries[folder.lower() + '/skill.md'])
    display, icon, icon_status = {}, None, 'category'
    presentation_files = None
    sidecar = entries.get((folder+'/'+SIDECAR).lower())
    if sidecar:
        try:
            raw_display = blob(sidecar,8192)
            display = presentation(raw_display)
            display_data = {SIDECAR:raw_display}
            for icon_name in set(icon_paths(display).values()):
                image_entry = entries.get((folder+'/'+icon_name).lower())
                if not image_entry or image_entry['path']!=folder+'/'+icon_name:
                    raise ValueError('icon_missing')
                display_data[icon_name] = blob(image_entry,MAX_ICON_BYTES)
            presentation_files = encode_files(display_data)
            preview = validate_files(presentation_files)
            icon, icon_status = preview['icon'], preview['icon_status']
        except (ValueError,TypeError,KeyError):
            display, icon, icon_status, presentation_files = {}, None, 'fallback', None
    metadata, body = parse_public(original)
    blockers = []
    if plugin_context and plugin_context['requirements']:
        blockers.append('插件附带能力尚未接入：'+', '.join(plugin_context['requirements']))
    if unsafe:
        blockers.append('不支持的文件类型：'+', '.join(unsafe[:8]))
    resources = {}
    reference_error = None
    presentation_paths = {folder+'/'+name for name in presentation_files or {}}
    unsupported = [path for path in extras if path not in presentation_paths and (PurePosixPath(path).suffix.lower() not in ('.md', '.txt') or path in unsafe)]
    if unsupported:
        blockers.append('资源/脚本尚未接通：' + ', '.join(unsupported[:8]))
    for path in extras:
        if path in unsupported or path in presentation_paths:
            continue
        data = blob(entries[path.lower()], MAX_REFERENCE_BYTES)
        resources[path[len(folder)+1:]] = {'text':data.decode('utf-8'), 'sha256':hashlib.sha256(data).hexdigest()}
    try:
        expanded = expand(body, resources)
    except ValueError as exc:
        reference_error = str(exc)
        expanded = body
        blockers.append('包内参考未通过检查：' + str(exc))
    if metadata.get('allowed-tools') or metadata.get('compatibility'):
        blockers.append('声明了尚未验证的工具或环境：' + metadata.get('allowed-tools', metadata.get('compatibility','')))
    if re.search(r'(?i)\b(sub.?agents?|python(?:-pptx)?|bash|powershell|npm|npx|mcp(?:__[a-z0-9_]+)?|oauth|skill tool)\b', body+'\n'+'\n'.join(item['text'] for item in resources.values())):
        blockers.append('正文或参考资料包含外部工具，当前不支持完整执行')
    license_items = []
    for parent in [PurePosixPath(folder), *PurePosixPath(folder).parents]:
        prefix = '' if str(parent) == '.' else str(parent) + '/'
        found = [item for item in entries.values() if item['path'].startswith(prefix) and '/' not in item['path'][len(prefix):] and PurePosixPath(item['path']).name in LICENSES]
        if found:
            license_items = found
            break
    license_items = list({item['path']:item for item in [*license_items, *[item for item in package if PurePosixPath(item['path']).name in LICENSES]]}.values())
    if not license_items:
        blockers.append('未找到适用许可文件')
    license_files = {item['path']:blob(item, 20000) for item in license_items}
    license_text = '\n\n'.join(path + '\n' + data.decode('utf-8') for path,data in sorted(license_files.items()))
    identifiers = {'MIT' if 'MIT License' in data.decode('utf-8') and 'Permission is hereby granted' in data.decode('utf-8') else 'Apache-2.0' if 'Apache License' in data.decode('utf-8') and 'Version 2.0' in data.decode('utf-8') else '待人工核对' for path,data in license_files.items() if PurePosixPath(path).name.startswith('LICENSE')}
    license_id = next(iter(identifiers)) if len(identifiers)==1 else '待人工核对'
    if license_id == '待人工核对':
        blockers.append('许可未识别，需人工核对后适配')
    declared_license = metadata.get('license') or (plugin_context or {}).get('declared_license')
    file_license = declared_license and re.fullmatch(r'Complete terms in (LICENSE(?:\.txt|\.md)?)', declared_license)
    if declared_license and declared_license != license_id and not (file_license and folder+'/'+file_license[1] in license_files):
        blockers.append('声明许可与已识别许可不一致，需人工核对')
    manifest = {'SKILL.md':hashlib.sha256(original).hexdigest(),
                **{name:item['sha256'] for name,item in resources.items()},
                **{name:item['sha256'] for name,item in (presentation_files or {}).items()},
                **{path[len(folder)+1:]:'git-blob-sha1:'+entries[path.lower()]['sha'] for path in unsupported},
                **{'@license/'+path:hashlib.sha256(data).hexdigest() for path,data in license_files.items()}}
    origin = {'repository':repository,'path':folder,'commit':commit,'source_sha256':hashlib.sha256(original).hexdigest(),
              'original_body_sha256':hashlib.sha256(body.encode()).hexdigest(),'license':license_id,
              'url':'https://github.com/' + repository + '/blob/' + commit + '/' + quote(folder, safe='/') + '/SKILL.md',
              'scope':'text-reference-only', 'manifest':manifest,
              'declarations':{key:value for key,value in metadata.items() if key not in ('name','description')},
              'original_description':metadata['description']}
    if plugin_context:
        origin['plugin_context'] = plugin_context
    content = None
    if not blockers:
        try:
            content = pack(metadata, body, origin, original, license_text, resources if resources else None,presentation_files=presentation_files)
        except ValueError as exc:
            blockers.append('封装超出限制：'+str(exc))
    slug = 'public-' + hashlib.sha256((repository + '/' + folder).encode()).hexdigest()[:12] + '-' + metadata['name'][:38].rstrip('-')
    if content is not None and len('按需技能 user:'+slug+'\n')+len(prompt_body(body,{**origin,'resources':resources})) > 8192:
        blockers.append('正文、参考与来源信息超出 Agent 加载预算')
        content = None
    progressive_content = None
    if resources and reference_error=='body_budget' and len(blockers)==1:
        try:
            progressive_content = pack(metadata,body,origin,original,license_text,resources,progressive=True,presentation_files=presentation_files)
        except ValueError:
            pass
    return {'name':metadata['name'],'description':metadata['description'],'origin':origin,'body':body,
            'license_text':license_text,'blockers':blockers,'compatible':not blockers,
            'resources':resources, 'manifest':manifest,
            'display':display, 'icon':icon, 'icon_status':icon_status,
            'progressive_available':progressive_content is not None,
            'compatibility_note':'仅作文字参考；选中技能时一并加载包内文字，不执行脚本或增加权限。',
            '_content':content,'_progressive_content':progressive_content,'slug':slug}
