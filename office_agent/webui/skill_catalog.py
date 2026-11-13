import hashlib
import json
from pathlib import Path
import re
import stat

from .public_skill_resolver import github_source


def featured():
    try:
        path = Path(__file__).resolve().parents[1]/'skill_catalog.json'
        info = path.lstat()
        if not stat.S_ISREG(info.st_mode) or path.is_symlink() or getattr(info,'st_file_attributes',0)&0x400 or info.st_size>32768:
            raise ValueError()
        with path.open('rb') as handle:
            data = handle.read(32769)
        if len(data)>32768:
            raise ValueError()
        catalog = json.loads(data)
        if catalog.get('schema')!=1 or not isinstance(catalog.get('items'),list) or not 1<=len(catalog['items'])<=8:
            raise ValueError()
        items, seen = [], set()
        for record in catalog['items']:
            repository,commit,folder = github_source(record['url'])
            identity = repository+'/'+folder+'@'+commit
            if identity in seen or not re.fullmatch('[0-9a-f]{64}',record['source_sha256']):
                raise ValueError()
            seen.add(identity)
            for key in ('name','display_name','description','category','license','reason'):
                if not isinstance(record[key],str) or not record[key].strip() or len(record[key])>300 or any(ord(c)<32 for c in record[key]):
                    raise ValueError()
            items.append({**record,'candidate_id':hashlib.sha256(('featured:'+identity).encode()).hexdigest()[:32],
                          'source':repository,'path':folder,'skill_id':folder.split('/')[-1],'commit':commit,
                          'provider':'featured','compatibility':'not_checked'})
        return {'ok':True,'items':items,'provider':'LiX 精选目录','reviewed_at':catalog['reviewed_at'],
                'snapshot_sha256':hashlib.sha256(data).hexdigest(),'network_requests':0}
    except (OSError,ValueError,KeyError,TypeError):
        return {'ok':False,'reason':'catalog_invalid','items':[]}
