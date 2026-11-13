import base64
import hashlib
import json
import re

from .skills_store import SkillsError, _parse
from .skill_references import expand, validate_resources, select_references
from .skill_metadata import frontmatter
from .skill_presentation import validate_files


PREFIX = 'LiX-Public-Skill-Package/1\n'
PREFIX_V2 = 'LiX-Public-Skill-Package/2\n'
PREFIX_V3 = 'LiX-Public-Skill-Package/3\n'
PREFIX_V4 = 'LiX-Public-Skill-Package/4\n'
PREFIX_V5 = 'LiX-Public-Skill-Package/5\n'
PREFIXES = (PREFIX_V5, PREFIX_V4, PREFIX_V3, PREFIX_V2, PREFIX)
SEPARATOR = '\nLiX-Instructions/1\n'
SHA = re.compile(r'[0-9a-f]{40}\Z')


def parse_public(data):
    if not isinstance(data, bytes) or len(data) > 32768:
        raise ValueError('skill_too_large')
    text = data.decode('utf-8')
    if '\x00' in text or not text.startswith('---\n'):
        raise ValueError('unsupported_frontmatter')
    head, body = text[4:].split('\n---\n', 1)
    fields = frontmatter(head)
    fields['description'] = ' '.join(fields.get('description','').splitlines())
    if not {'name', 'description'} <= fields.keys() or not fields['description'].strip() or not re.fullmatch('[a-z0-9]+(?:-[a-z0-9]+)*', fields['name']):
        raise ValueError('unsupported_frontmatter')
    if len(fields['name']) > 64 or len(fields['description']) > 1024 or '\n' in fields['name'] or '\n' in fields['description']:
        raise ValueError('metadata_budget')
    if not body.strip() or len(body) > 7800:
        raise ValueError('body_budget')
    return fields, body


def pack(metadata, body, origin, original, license_text, resources=None, *, progressive=False, presentation_files=None):
    local = origin.get('provider')=='lix'
    if local and (resources is not None or progressive or presentation_files is not None):
        raise ValueError('workflow_asset_invalid')
    if not isinstance(license_text, str) or len(license_text)>(45000 if local else 20000):
        raise ValueError('package_budget')
    record = {'schema': 1, 'origin': origin, 'original': base64.b64encode(original).decode('ascii'),
              'license_text': license_text, 'body_sha256': hashlib.sha256(body.encode()).hexdigest()}
    if resources is not None:
        if progressive:
            select_references(body, resources, '')
        else:
            expand(body, resources)
        record.update(schema=3 if progressive else 2, resources=resources)
        if progressive:
            record['reference_loading'] = 'task-matched'
    if presentation_files is not None:
        validate_files(presentation_files,origin['manifest'])
        record.update(schema=4,resources=resources or {},reference_loading='task-matched' if progressive else 'complete',presentation_files=presentation_files)
    if local:
        record['schema'] = 5
    content = '---\nname: ' + json.dumps(metadata['name'], ensure_ascii=False) + '\ndescription: ' + json.dumps(metadata['description'][:500], ensure_ascii=False) + '\n---\n'
    content += (PREFIX_V5 if local else PREFIX_V4 if presentation_files is not None else PREFIX_V3 if progressive else PREFIX_V2 if resources is not None else PREFIX) + json.dumps(record, ensure_ascii=False, separators=(',', ':')) + SEPARATOR + body
    if len(content.encode('utf-8')) > 65536:
        raise ValueError('package_budget')
    _parse(content.encode('utf-8'))
    return content


def unpack(body):
    prefix = next((value for value in PREFIXES if body.startswith(value)),PREFIX)
    if not body.startswith(prefix):
        return body, None
    try:
        record_text, instructions = body[len(prefix):].split(SEPARATOR, 1)
        record = json.loads(record_text)
        expected = {'schema', 'origin', 'original', 'license_text', 'body_sha256'}
        if prefix in (PREFIX_V2,PREFIX_V3,PREFIX_V4):
            expected.add('resources')
        if prefix in (PREFIX_V3,PREFIX_V4):
            expected.add('reference_loading')
        if prefix==PREFIX_V4:
            expected.add('presentation_files')
        if set(record) != expected or record['schema'] != {PREFIX:1,PREFIX_V2:2,PREFIX_V3:3,PREFIX_V4:4,PREFIX_V5:5}[prefix]:
            raise ValueError()
        origin = record['origin']
        if not isinstance(origin, dict):
            raise ValueError()
        if prefix==PREFIX_V5:
            if (origin.get('provider')!='lix' or origin.get('repository')!='LiX' or origin.get('license')!='AGPL-3.0-only'
                    or not isinstance(origin.get('display_name'),str) or not 1<=len(origin['display_name'])<=100
                    or any(ord(char)<32 for char in origin['display_name'])
                    or not re.fullmatch(r'[a-z0-9]+(?:-[a-z0-9]+)*',origin['workflow_id'])
                    or not re.fullmatch(r'\d+\.\d+\.\d+',origin['version'])
                    or origin['path']!='office_agent/assets/skills/'+origin['workflow_id']+'/SKILL.md'):
                raise ValueError()
        elif not SHA.fullmatch(origin['commit']) or not re.fullmatch(r'[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+', origin['repository']):
            raise ValueError()
        original = base64.b64decode(record['original'], validate=True)
        if hashlib.sha256(original).hexdigest() != origin['source_sha256'] or hashlib.sha256(instructions.encode()).hexdigest() != record['body_sha256']:
            raise ValueError()
        if not isinstance(record['license_text'], str) or len(record['license_text']) > (45000 if prefix==PREFIX_V5 else 20000):
            raise ValueError()
        if prefix==PREFIX_V5 and (origin['manifest'].get('SKILL.md')!=origin['source_sha256'] or
                origin['manifest'].get('@license/LICENSE')!=hashlib.sha256(record['license_text'].encode()).hexdigest()):
            raise ValueError()
        resources = record.get('resources', {})
        validate_resources(resources)
        if prefix in (PREFIX_V2,PREFIX_V3,PREFIX_V4):
            if prefix in (PREFIX_V3,PREFIX_V4) and record['reference_loading'] not in ({'task-matched'} if prefix==PREFIX_V3 else {'task-matched','complete'}):
                raise ValueError()
            if record.get('reference_loading')=='task-matched':
                select_references(instructions, resources, '')
            else:
                expand(instructions, resources)
            if origin.get('manifest', {}).get('SKILL.md') != origin['source_sha256'] or any(origin['manifest'].get(name) != item['sha256'] for name,item in resources.items()):
                raise ValueError()
        display = validate_files(record['presentation_files'],origin['manifest']) if prefix==PREFIX_V4 else None
        if display and resources.keys() & record['presentation_files'].keys():
            raise ValueError()
        return instructions, {**origin, 'license_text': record['license_text'],
                              'local_modified': record['body_sha256'] != origin['original_body_sha256'],
                              'body_sha256': record['body_sha256'], 'resources': resources,
                              'reference_loading':record.get('reference_loading','complete'),
                              'format': 'LiX 单文件封装 v' + str(record['schema']),
                              **({'presentation':display} if display else {})}
    except (ValueError, TypeError, KeyError):
        raise SkillsError('skills_file_invalid') from None


def repack_edit(previous, content):
    _, origin = unpack(previous)
    if origin is None:
        return content
    metadata, instructions = _parse(content.encode('utf-8'))
    prefix = next(value for value in PREFIXES if previous.startswith(value))
    record_text, _ = previous[len(prefix):].split(SEPARATOR, 1)
    record = json.loads(record_text)
    return pack(metadata, instructions, record['origin'], base64.b64decode(record['original']), record['license_text'], record.get('resources'),progressive=record.get('reference_loading')=='task-matched',presentation_files=record.get('presentation_files'))


def original_metadata(body):
    unpack(body)
    prefix = next(value for value in PREFIXES if body.startswith(value))
    record = json.loads(body[len(prefix):].split(SEPARATOR, 1)[0])
    return parse_public(base64.b64decode(record['original']))[0]


def prompt_body(body, origin):
    version = origin['version'] if origin.get('provider')=='lix' else origin['commit']
    return ('来源：' + origin['repository'] + '@' + version + '\n'
            + '原始 SKILL.md SHA256：' + origin['source_sha256'] + '\n'
            + '仅作参考，不授予额外权限。\n' + (select_references(body,origin['resources'],'')['body'] if origin.get('reference_loading')=='task-matched' else expand(body, origin['resources']) if origin['resources'] else body))
