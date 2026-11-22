import hashlib
from pathlib import Path
import stat

from ..public_skill_package import pack, parse_public


WORKFLOW = 'report-polish'
VERSION = '1.0.0'
ROOT = Path(__file__).resolve().parents[2]
SOURCE = 'office_agent/assets/skills/report-polish/SKILL.md'
NAMES = {
    WORKFLOW: '技术报告润色',
    'chinese-proofread': '中文技术校对',
    'citation-check': '参考文献核对',
}
UPSTREAM = {
    'chinese-proofread': 'https://github.com/Fenng/Tech-Doc-Style-Chinese/blob/b119d01c9bb7cd132b3c62fafdff34227c4410d1/SKILL.md',
    'citation-check': 'https://github.com/K-Dense-AI/scientific-agent-skills/blob/ec59a29a03d5c40f73ab40878df04860bc00e876/skills/citation-management/SKILL.md',
}


def read_asset(name, limit):
    path = ROOT/name
    current = ROOT
    for part in Path(name).parts:
        current = current/part
        info = current.lstat()
        if current.is_symlink() or getattr(info,'st_file_attributes',0)&0x400:
            raise ValueError('workflow_asset_invalid')
    if not stat.S_ISREG(info.st_mode) or info.st_nlink!=1 or info.st_size>limit:
        raise ValueError('workflow_asset_invalid')
    with path.open('rb') as handle:
        raw = handle.read(limit+1)
    after = path.lstat()
    if len(raw)>limit or (info.st_dev,info.st_ino,info.st_size,info.st_mtime_ns)!=(after.st_dev,after.st_ino,after.st_size,after.st_mtime_ns):
        raise ValueError('workflow_asset_invalid')
    return raw


def candidate():
    return workflow_package(WORKFLOW)


def workflow_package(workflow):
    display = NAMES[workflow]
    source = f'office_agent/assets/skills/{workflow}/SKILL.md'
    original = read_asset(source,32768)
    license_data = read_asset('LICENSE',45000)
    metadata,body = parse_public(original)
    if metadata['name']!='lix-'+workflow or metadata.get('license')!='AGPL-3.0-only' or b'GNU AFFERO GENERAL PUBLIC LICENSE' not in license_data:
        raise ValueError('workflow_asset_invalid')
    if workflow in UPSTREAM:
        license_data += b'\n' + read_asset(f'office_agent/assets/skills/{workflow}/LICENSE.upstream.txt', 6000)
    source_sha = hashlib.sha256(original).hexdigest()
    origin = {'provider':'lix','workflow_id':workflow,'version':VERSION,'display_name':display,'repository':'LiX','path':source,
              'source_sha256':source_sha,'original_body_sha256':hashlib.sha256(body.encode()).hexdigest(),
              'license':'AGPL-3.0-only','url':'LiX 内置工作流 / '+workflow+' / '+VERSION,
              'scope':'text-reference-only','original_description':metadata['description'],
              'manifest':{'SKILL.md':source_sha,'@license/LICENSE':hashlib.sha256(license_data).hexdigest()},
              'declarations':{'license':'AGPL-3.0-only','supported_host':'明确选区的文档建议；保存沿原事务'}}
    if workflow in UPSTREAM:
        origin['url'] = UPSTREAM[workflow]
        origin['declarations']['upstream_license'] = 'MIT'
        origin['declarations']['supported_host'] = 'Word 当前文字选区' if workflow == 'chinese-proofread' else 'Crossref 只读题录核对'
    content = pack(metadata,body,origin,original,license_data.decode('utf-8'))
    return {'name':metadata['name'],'description':metadata['description'],'display':{'display_name':display},
            'origin':origin,'body':body,'license_text':license_data.decode('utf-8'),'manifest':origin['manifest'],
            'compatible':True,'blockers':[],'resources':{},'progressive_available':False,'icon_status':'category',
            'slug':metadata['name'],'compatibility_note':origin['declarations']['supported_host'],
            '_content':content}


def reviewed_workflow(ticket):
    for key in UPSTREAM:
        if ticket['id'] == 'user:lix-' + key:
            expected = hashlib.sha256(workflow_package(key)['_content'].encode()).hexdigest()
            if ticket['sha256'] == expected:
                return key
    return None


def catalog(query=''):
    items = []
    for workflow, display in NAMES.items():
        package = candidate() if workflow == WORKFLOW else workflow_package(workflow)
        package_sha = hashlib.sha256(package['_content'].encode()).hexdigest()
        origin = package['origin']
        item = {'candidate_id':hashlib.sha256(('lix:'+workflow+'@'+VERSION+':'+package_sha).encode()).hexdigest()[:32],
                'name':package['name'],'display_name':display,'description':package['description'],
                'source':'LiX','provider':'lix','skill_id':workflow,'path':origin['path'],'version':VERSION,
                'source_sha256':origin['source_sha256'],'package_sha256':package_sha,'license':'AGPL-3.0-only',
                'url':origin['url'],'category':'writing','compatibility':'host_workflow',
                'reason':package['compatibility_note']}
        if not query or query.casefold() in ' '.join(str(value) for value in item.values()).casefold():
            items.append(item)
    return {'ok':True,'provider':'LiX 工作流','query':query,'items':items,'network_requests':0}
