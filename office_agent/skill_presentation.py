import base64
import hashlib
from io import BytesIO
from pathlib import PurePosixPath

from PIL import Image

from .skill_metadata import presentation
from .skill_references import safe_name


MAX_ICON_BYTES = 512 * 1024
MAX_ICON_SIDE = 1024
SIDECAR = 'agents/openai.yaml'
IMAGE_SUFFIXES = {'.png':'PNG', '.jpg':'JPEG', '.jpeg':'JPEG', '.webp':'WEBP'}


def thumbnail(data, *, filename=None):
    if not isinstance(data, bytes) or len(data)>MAX_ICON_BYTES:
        raise ValueError('icon_budget')
    try:
        with Image.open(BytesIO(data)) as source:
            if source.format not in set(IMAGE_SUFFIXES.values()) or max(source.size)>MAX_ICON_SIDE or min(source.size)<1 or getattr(source,'n_frames',1)!=1:
                raise ValueError('icon_format')
            if filename is not None and IMAGE_SUFFIXES.get(PurePosixPath(filename).suffix.lower())!=source.format:
                raise ValueError('icon_format')
            source.load()
            cleaned = source.convert('RGBA')
            cleaned.thumbnail((96,96),Image.Resampling.LANCZOS)
            output = BytesIO()
            cleaned.save(output,format='PNG')
        return {'data_url':'data:image/png;base64,'+base64.b64encode(output.getvalue()).decode('ascii'),
                'source_sha256':hashlib.sha256(data).hexdigest(),'width':cleaned.width,'height':cleaned.height}
    except Exception:
        raise ValueError('icon_invalid') from None


def icon_paths(display):
    paths = {}
    for key in ('icon_small','icon_large'):
        if key not in display:
            continue
        path = display[key].removeprefix('./')
        safe_name(path)
        if PurePosixPath(path).suffix.lower() not in IMAGE_SUFFIXES:
            raise ValueError('icon_format')
        paths[key] = path
    if len({path.lower() for path in paths.values()})!=len(set(paths.values())):
        raise ValueError('unsupported_display_metadata')
    return paths


def validate_files(files, manifest=None):
    if not isinstance(files,dict) or not 1<=len(files)<=3 or SIDECAR not in files:
        raise ValueError('unsupported_display_metadata')
    decoded = {}
    for path,item in files.items():
        safe_name(path)
        if not isinstance(item,dict) or set(item)!={'base64','sha256'} or not isinstance(item['base64'],str) or len(item['base64'])>MAX_ICON_BYTES*2:
            raise ValueError('unsupported_display_metadata')
        data = base64.b64decode(item['base64'],validate=True)
        if hashlib.sha256(data).hexdigest()!=item['sha256'] or (manifest is not None and manifest.get(path)!=item['sha256']):
            raise ValueError('source_integrity')
        decoded[path] = data
    display = presentation(decoded[SIDECAR])
    paths = icon_paths(display)
    if set(files)!={SIDECAR,*paths.values()}:
        raise ValueError('unsupported_display_metadata')
    icons = {path:thumbnail(decoded[path],filename=path) for path in set(paths.values())}
    icon = icons.get(paths.get('icon_small') or paths.get('icon_large'))
    return {'display':display,'icon':icon,'icon_status':'verified-source' if icon else 'category',
            'files':[{'path':path,'sha256':files[path]['sha256'],'size':len(data)} for path,data in sorted(decoded.items())]}


def encode_files(data):
    files = {path:{'base64':base64.b64encode(raw).decode('ascii'),'sha256':hashlib.sha256(raw).hexdigest()} for path,raw in data.items()}
    validate_files(files)
    return files
