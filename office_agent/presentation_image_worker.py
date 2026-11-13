from io import BytesIO
import hashlib
import json
from pathlib import Path
import sys
import warnings


def encode(image, format):
    output=BytesIO()
    image.save(output,format=format,**({'quality':90,'optimize':True} if format=='JPEG' else {}))
    raw=output.getvalue()
    from PIL import Image
    with Image.open(BytesIO(raw)) as check:
        check.load()
        if set(check.info)-{'jfif','jfif_version','jfif_unit','jfif_density'} or check.size!=image.size or check.format!=format:raise ValueError('metadata')
    return raw


def process(root):
    from PIL import Image, ImageOps
    Image.MAX_IMAGE_PIXELS=40_000_000
    warnings.simplefilter('error',Image.DecompressionBombWarning)
    raw=(root/'input').read_bytes()
    if not 0<len(raw)<=10_000_000:raise ValueError('capacity')
    with Image.open(BytesIO(raw)) as source:
        format=source.format
        size=source.size
        if format not in {'PNG','JPEG'} or getattr(source,'n_frames',1)!=1 or max(size)>32768 or size[0]*size[1]>40_000_000:raise ValueError('format')
        source.verify()
    with Image.open(BytesIO(raw)) as source:
        source.load()
        oriented=ImageOps.exif_transpose(source)
        try:
            alpha=oriented.mode in {'RGBA','LA'} or 'transparency' in oriented.info
            converted=oriented.convert('RGBA' if alpha else 'RGB')
            try:
                clean=Image.new(converted.mode,converted.size);clean.paste(converted)
            finally:converted.close()
        finally:
            if oriented is not source:oriented.close()
    try:
        normalized_format='PNG' if alpha or format=='PNG' else 'JPEG'
        base=encode(clean,normalized_format)
        if len(base)>10_000_000:base=None
        resized=clean.copy();resized.thumbnail((1600,1600),Image.Resampling.LANCZOS)
        try:
            optimized_format='PNG' if alpha else 'JPEG'
            optimized=encode(resized,optimized_format)
            thumb=clean.copy();thumb.thumbnail((260,180),Image.Resampling.LANCZOS)
            try:thumbnail=encode(thumb,'PNG')
            finally:thumb.close()
            report={'original':{'format':format,'width':size[0],'height':size[1],'bytes':len(raw)},'variants':{}}
            for name,data,dimensions,kind in [('normalized',base,clean.size,normalized_format),('optimized',optimized,resized.size,optimized_format)]:
                if data is not None:
                    (root/name).write_bytes(data)
                    report['variants'][name]={'sha256':hashlib.sha256(data).hexdigest(),'bytes':len(data),'width':dimensions[0],'height':dimensions[1],'format':kind}
                    with Image.open(BytesIO(data)) as preview:
                        preview.thumbnail((260,180),Image.Resampling.LANCZOS)
                        copy=Image.new(preview.mode,preview.size);copy.paste(preview)
                        try:(root/(name+'-thumbnail')).write_bytes(encode(copy,'PNG'))
                        finally:copy.close()
            (root/'thumbnail').write_bytes(thumbnail)
            (root/'report.json').write_text(json.dumps(report),encoding='utf8')
        finally:resized.close()
    finally:clean.close()


if __name__=='__main__':
    try:process(Path(sys.argv[1]))
    except Exception:raise SystemExit(1)
