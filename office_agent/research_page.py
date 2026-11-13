import base64
from contextlib import closing
import io
import json
import math
from pathlib import Path
import sys
import tempfile

from .fidelity import _run, _temporary_root
from .rag_sources import MAX_SOURCE_BYTES, SourceError
from .writes import digest


def render_payload(path, *, page, box):
    import pypdfium2 as pdfium
    if type(page) is not int or not 0 <= page < 2000 or type(box) is not list or len(box) != 4 or any(type(value) not in (int, float) or not math.isfinite(value) for value in box):
        raise SourceError('invalid')
    with pdfium.PdfDocument(str(path)) as document:
        if not 1 <= len(document) <= 2000 or page >= len(document):
            raise SourceError('capacity')
        with closing(document[page]) as current:
            width, height = current.get_size()
            if not all(math.isfinite(value) and 0 < value < 1_000_000 for value in (width, height)):
                raise SourceError('capacity')
            scale = min(1.5, 1000 / max(width, height))
            with closing(current.render(scale=scale)) as bitmap, bitmap.to_pil() as picture:
                buffer = io.BytesIO()
                with picture.convert('RGB') as rgb:
                    rgb.save(buffer, format='JPEG', quality=65)
                image = buffer.getvalue()
                if len(image) > 120000:
                    raise SourceError('capacity')
                coordinates = None
                try:
                    converter = bitmap.get_posconv(current)
                    points = [converter.to_bitmap(x, y) for x in (box[0], box[2]) for y in (box[1], box[3])]
                    left, right = min(point[0] for point in points), max(point[0] for point in points)
                    top, bottom = min(point[1] for point in points), max(point[1] for point in points)
                    if 0 <= left < right <= picture.width and 0 <= top < bottom <= picture.height:
                        coordinates = [left, top, right - left, bottom - top]
                except Exception:
                    pass
                return {'image': 'data:image/jpeg;base64,' + base64.b64encode(image).decode('ascii'), 'width': picture.width, 'height': picture.height,
                        'highlight': coordinates, 'precision': 'pdfium_page_to_device_run_bounds' if coordinates else 'page_only',
                        'page': page + 1, 'rotation': current.get_rotation(), 'crop_box': list(current.get_cropbox()),
                        'notice': '80字符合并边界，不是逐词定位；多栏顺序及文字正确性需人工核对。' if coordinates else '坐标无法可靠映射，仅定位到原页。'}


def render_page(path, fingerprint, page, box, *, ocr=False):
    path = Path(path)
    if path.stat().st_size > MAX_SOURCE_BYTES or digest(path) != fingerprint:
        raise SourceError('source_changed')
    processes = []
    with tempfile.TemporaryDirectory(prefix='office-page-', dir=_temporary_root()) as raw:
        scratch = Path(raw)
        source, options, target = scratch / 'input.pdf', scratch / 'options.json', scratch / 'result.json'
        source.write_bytes(path.read_bytes())
        if digest(source) != fingerprint:
            raise SourceError('source_changed')
        options.write_text(json.dumps({'page': page} if ocr else {'page': page, 'box': box}), encoding='utf8')
        _run((sys.executable, str(Path(__file__).with_name('read_worker.py')), 'pdf-ocr' if ocr else 'pdf-page', str(source), str(options), str(target)), 10000, processes)
        if not target.is_file() or target.stat().st_size > 180000:
            raise SourceError('capacity')
        result = json.loads(target.read_bytes())
    if result.get('error'):
        raise SourceError(result['error'])
    if digest(path) != fingerprint:
        raise SourceError('source_changed')
    return {**result, 'source_sha256': fingerprint, 'processes': processes, 'read_only': True}
