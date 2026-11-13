import hashlib
import json
from pathlib import Path

from PIL import Image


ROOT = Path(__file__).resolve().parents[1]
ASSETS = ROOT / 'office_agent/webui/static/brand'
SIZES = (16, 24, 32, 48, 64, 128, 256)


def build():
    source = ASSETS / 'lix-source.png'
    transparent = ASSETS / 'lix-transparent.png'
    with Image.open(transparent) as image:
        image.load()
        if image.width != image.height or image.width > 4096:
            raise ValueError('Expected a square brand image up to 4096 px')
        rgba = image.convert('RGBA')
        rgba.resize((256, 256), Image.Resampling.LANCZOS).save(ASSETS / 'lix.png')
        frames = [rgba.resize((size, size), Image.Resampling.HAMMING) for size in SIZES]
        frames[-1].save(ASSETS / 'lix.ico', sizes=[(size, size) for size in SIZES], append_images=frames[:-1])
    manifest = {
        'name': 'LiX', 'source': 'User-supplied artwork, 2026-11-12',
        'transform': 'Original retained; imagegen orange background with enlarged artwork and transparent exterior, 2026-11-13; PNG Lanczos, ICO Hamming to avoid alpha ringing',
        'icon_sizes': list(SIZES),
        'files': {path.name: hashlib.sha256(path.read_bytes()).hexdigest()
                  for path in (source, transparent, ASSETS / 'lix.png', ASSETS / 'lix.ico')},
    }
    (ASSETS / 'manifest.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')


if __name__ == '__main__':
    build()
