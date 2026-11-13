import asyncio
from contextlib import closing
import math
import platform
from importlib.metadata import version

from .rag_sources import SourceError


async def recognize(picture):
    from winrt.windows.graphics.imaging import SoftwareBitmap, BitmapPixelFormat
    from winrt.windows.media.ocr import OcrEngine
    from winrt.windows.storage.streams import Buffer

    engine = OcrEngine.try_create_from_user_profile_languages()
    if engine is None:
        raise SourceError('ocr_unavailable')
    pixels = picture.convert('L')
    try:
        content = pixels.tobytes()
        buffer = Buffer(len(content))
        buffer.length = len(content)
        memoryview(buffer)[:] = content
        bitmap = SoftwareBitmap.create_copy_from_buffer(buffer, BitmapPixelFormat.GRAY8, picture.width, picture.height)
        try:
            result = await engine.recognize_async(bitmap)
            lines = []
            for line in result.lines:
                words = list(line.words)
                if not words:
                    continue
                rectangles = [word.bounding_rect for word in words]
                left, top = min(row.x for row in rectangles), min(row.y for row in rectangles)
                right, bottom = max(row.x + row.width for row in rectangles), max(row.y + row.height for row in rectangles)
                lines.append({'text': line.text, 'box_pixels': [left, top, right, bottom]})
            if len(result.text) > 12000 or len(lines) > 120:
                raise SourceError('capacity')
            return {'raw_text': result.text, 'lines': lines, 'language': engine.recognizer_language.language_tag,
                    'engine': 'Windows.Media.Ocr', 'engine_version': platform.version() + ':' + version('winrt-Windows.Media.Ocr'),
                    'quality': 'ocr_unverified', 'numeric_reliability': 'not_verified',
                    'warning': '原始OCR可能混淆数字、单位、上下标、符号及多栏顺序。人工校正单独保存，不替换原始OCR。'}
        finally:
            bitmap.close()
    finally:
        pixels.close()


def ocr_payload(path, *, page):
    import pypdfium2 as pdfium
    if type(page) is not int or not 0 <= page <= 1999:
        raise SourceError('invalid')
    try:
        with pdfium.PdfDocument(str(path)) as document:
            if page >= len(document) or len(document) > 2000:
                raise SourceError('invalid')
            with closing(document[page]) as current:
                width, height = current.get_size()
                if not all(math.isfinite(value) and 0 < value < 1_000_000 for value in (width, height)):
                    raise SourceError('capacity')
                with closing(current.render(scale=min(2, 2000 / max(width, height)))) as bitmap, bitmap.to_pil() as picture:
                    result = asyncio.run(recognize(picture))
                    converter = bitmap.get_posconv(current)
                    for line in result['lines']:
                        left, top, right, bottom = line['box_pixels']
                        points = [converter.to_page(round(x), round(y)) for x in (left, right) for y in (top, bottom)]
                        line['box_pdf'] = [min(point[0] for point in points), min(point[1] for point in points), max(point[0] for point in points), max(point[1] for point in points)]
                    return result
    except ImportError:
        raise SourceError('ocr_unavailable') from None
    except SourceError:
        raise
    except Exception:
        raise SourceError('unsupported') from None
