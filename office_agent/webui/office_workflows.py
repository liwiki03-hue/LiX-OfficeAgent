import hashlib
from pathlib import Path
import stat

from ..public_skill_package import pack, parse_public


WORKFLOW = 'report-polish'
VERSION = '1.0.0'
ROOT = Path(__file__).resolve().parents[2]
SOURCE = 'office_agent/assets/skills/report-polish/SKILL.md'


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
    original = read_asset(SOURCE,32768)
    license_data = read_asset('LICENSE',45000)
    metadata,body = parse_public(original)
    if metadata['name']!='lix-report-polish' or metadata.get('license')!='AGPL-3.0-only' or b'GNU AFFERO GENERAL PUBLIC LICENSE' not in license_data:
        raise ValueError('workflow_asset_invalid')
    source_sha = hashlib.sha256(original).hexdigest()
    origin = {'provider':'lix','workflow_id':WORKFLOW,'version':VERSION,'display_name':'技术报告润色','repository':'LiX','path':SOURCE,
              'source_sha256':source_sha,'original_body_sha256':hashlib.sha256(body.encode()).hexdigest(),
              'license':'AGPL-3.0-only','url':'LiX 内置工作流 / '+WORKFLOW+' / '+VERSION,
              'scope':'text-reference-only','original_description':metadata['description'],
              'manifest':{'SKILL.md':source_sha,'@license/LICENSE':hashlib.sha256(license_data).hexdigest()},
              'declarations':{'license':'AGPL-3.0-only','supported_host':'明确选区的文档建议；保存沿原事务'}}
    content = pack(metadata,body,origin,original,license_data.decode('utf-8'))
    return {'name':metadata['name'],'description':metadata['description'],'display':{'display_name':'技术报告润色'},
            'origin':origin,'body':body,'license_text':license_data.decode('utf-8'),'manifest':origin['manifest'],
            'compatible':True,'blockers':[],'resources':{},'progressive_available':False,'icon_status':'category',
            'slug':'lix-report-polish','compatibility_note':'LiX原创文字工作流；只提供参考，不新增工具或自动修改文件。',
            '_content':content}


def catalog(query=''):
    package = candidate()
    package_sha = hashlib.sha256(package['_content'].encode()).hexdigest()
    item = {'candidate_id':hashlib.sha256(('lix:'+WORKFLOW+'@'+VERSION+':'+package_sha).encode()).hexdigest()[:32],
            'name':package['name'],'display_name':'技术报告润色','description':package['description'],
            'source':'LiX','provider':'lix','skill_id':WORKFLOW,'path':SOURCE,'version':VERSION,
            'source_sha256':package['origin']['source_sha256'],'package_sha256':package_sha,'license':'AGPL-3.0-only',
            'url':package['origin']['url'],'category':'writing','compatibility':'host_workflow',
            'reason':'LiX提供；用于明确选区的受约束润色，安装与写入分别确认。'}
    matches = not query or query.casefold() in (' '.join(str(value) for value in item.values())).casefold()
    return {'ok':True,'provider':'LiX 工作流','query':query,'items':[item] if matches else [],'network_requests':0}
