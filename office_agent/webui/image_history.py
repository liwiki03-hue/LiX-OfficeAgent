import base64
import hashlib
import json
import os
from pathlib import Path
import re

from ..imagegen import ImageGenError, ImageRequest
from ..imageprovider import normalize_generated_image
from ..imagetransactions import MAX_OUTPUT_BYTES, _identity, _stamp
from ..rag_evidence import _read, _dir
from ..writes import _audit_path, Workspace


def history_action(owner, identifier, operation):
    with owner.lock:
        if not owner.authenticated():
            raise ImageGenError('not_authenticated')
        if type(identifier) is not str or not re.fullmatch('[a-f0-9-]{36}',identifier):
            raise ImageGenError('image_history_invalid')
        ticket = owner.image_tickets.get(identifier)
        if ticket is None or ticket[0] != owner.state.epoch:
            raise ImageGenError('image_history_expired')
        raw, receipt = _read(ticket[1])
        if receipt != ticket[2]:
            raise ImageGenError('image_history_changed')
        record = json.loads(raw)
        if record.get('source') is not None or record.get('transaction_id') != identifier or record.get('kind') != 'image_generation':
            raise ImageGenError('image_history_changed')
        if operation == 'reuse':
            value = record.get('reusable')
            if type(value) is not dict or set(value) != {'version','prompt','size','style','n'} or value['version'] != 1:
                raise ImageGenError('image_history_prompt_unavailable')
            request = ImageRequest(**{key:value[key] for key in ('prompt','size','style','n')})
            if hashlib.sha256(request.effective_prompt.encode('utf8')).hexdigest() != record.get('prompt_sha256'):
                raise ImageGenError('image_history_changed')
            from ..settings_store import _private_values
            owner._catalog()
            store = owner.settings._store
            if any(secret and secret in request.prompt for secret in _private_values(store._values(store._read(store.root / '.env',True)))):
                raise ImageGenError('image_history_sensitive')
            current = owner.capture('image')
            return {'ok':True,'draft':request.model_dump(),'route':current.public(),'scope':'仅填入草稿；当前接口重新解析，未恢复生成、保存或撤销批准。'}
        output = record.get('saved_output')
        if record.get('saved') is not True or type(output) is not dict:
            raise ImageGenError('image_history_not_saved')
        path = Path(output.get('path',''))
        if not path.is_absolute() or not re.fullmatch(r'office-image-[a-f0-9-]{36}\.png',path.name):
            raise ImageGenError('image_history_changed')
        if type(output.get('bytes')) is not int or not 0 < output['bytes'] <= MAX_OUTPUT_BYTES:
            raise ImageGenError('image_history_too_large')
        _audit_path(path.parent, directory=True)
        if not os.path.lexists(path):
            raise ImageGenError('image_history_missing')
        _audit_path(path, directory=False)
        space = Workspace(path.parent)
        if space.file(str(path)) != path:
            raise ImageGenError('image_history_changed')
        parent = _dir(path.parent)
        before = path.lstat()
        if before.st_size > MAX_OUTPUT_BYTES:
            raise ImageGenError('image_history_too_large')
        if before.st_nlink != 1 or before.st_size != output['bytes'] or ('file_identity' in output and list(_identity(before)) != output['file_identity']):
            raise ImageGenError('image_history_changed')
        with path.open('rb') as stream:
            first = os.fstat(stream.fileno())
            if _stamp(first)[:-1] != _stamp(before)[:-1] or first.st_nlink != 1:
                raise ImageGenError('image_history_changed')
            data = stream.read(MAX_OUTPUT_BYTES+1)
            last = os.fstat(stream.fileno())
            if _stamp(last) != _stamp(first) or last.st_nlink != 1:
                raise ImageGenError('image_history_changed')
        after = path.lstat()
        if _stamp(after) != _stamp(before) or after.st_nlink != 1 or _dir(path.parent) != parent or hashlib.sha256(data).hexdigest() != output.get('sha256') or _read(ticket[1])[1] != receipt:
            raise ImageGenError('image_history_changed')
        image = normalize_generated_image(data)
        if operation == 'open':
            os.startfile(str(path.parent))
            return {'ok':True,'path':str(path.parent)}
        if operation != 'preview':
            raise ImageGenError('image_history_invalid')
        return {'ok':True,'image':{'data_url':'data:image/png;base64,' + base64.b64encode(image.raw).decode('ascii'),
                                  'alt':path.name,'caption':'已核验保存记录的大小与摘要；' + ('文件身份匹配。' if 'file_identity' in output else '旧记录无原始文件身份，仅核对摘要及本次读取稳定性。')},
                'path':str(path),'verified_sha256':output['sha256'],'read_only':True}
