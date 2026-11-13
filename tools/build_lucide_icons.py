import argparse
import hashlib
import json
from pathlib import Path

from PySide6.QtCore import QByteArray, QRectF
from PySide6.QtGui import QImage, QPainter
from PySide6.QtSvg import QSvgRenderer


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--target', type=Path, required=True)
    args = parser.parse_args()
    root = Path(__file__).resolve().parents[1]
    icons = root / 'office_agent/webui/static/vendor/lucide'
    sources = json.loads((root / 'release/LUCIDE_SOURCE.json').read_text('utf-8'))
    expected = {item['path']: item['sha256'] for item in sources['files']}
    target = args.target.absolute()
    if target.exists() or not target.parent.is_dir():
        raise ValueError('Choose a new directory under an existing parent')
    for parent in target.parents:
        if parent.is_symlink() or getattr(parent.lstat(), 'st_file_attributes', 0) & 0x400:
            raise ValueError('Links are not allowed')
    target.mkdir()
    records = []
    for name in ('folder-open', 'message-square', 'history', 'settings', 'x'):
        source_name = 'rotate-ccw-clock' if name == 'history' else name
        data = (icons / (source_name + '.svg')).read_bytes()
        if hashlib.sha256(data).hexdigest() != expected[source_name + '.svg']:
            raise ValueError('Fixed icon source changed')
        renderer = QSvgRenderer(QByteArray(data.replace(b'currentColor', b'#f8faff')))
        if not renderer.isValid():
            raise ValueError('Invalid SVG')
        image = QImage(80, 80, QImage.Format.Format_ARGB32_Premultiplied)
        image.fill(0)
        painter = QPainter(image)
        try:
            renderer.render(painter, QRectF(0, 0, 80, 80))
        finally:
            painter.end()
        output = target / (name + '.png')
        if not image.save(str(output)):
            raise ValueError('PNG write failed')
        records.append({'file': output.name, 'sha256': hashlib.sha256(output.read_bytes()).hexdigest(),
                        'matches_vendored': output.read_bytes() == (icons / 'native' / output.name).read_bytes()})
    (target / 'RESULT.json').write_text(json.dumps(records, indent=2), encoding='utf-8')
    print(json.dumps({'icons': len(records), 'all_match': all(item['matches_vendored'] for item in records)}))


if __name__ == '__main__':
    main()
