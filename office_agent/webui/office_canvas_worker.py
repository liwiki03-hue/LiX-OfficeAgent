import json
from pathlib import Path
import sys


def main():
    import pypdfium2 as pdfium
    source, index, folder = Path(sys.argv[1]), int(sys.argv[2]), Path(sys.argv[3])
    with pdfium.PdfDocument(str(source)) as book:
        if not 1 <= len(book) <= 100 or not 0 <= index < len(book):
            return 2
        page = book[index]
        try:
            bitmap = page.render(scale=min(1400/page.get_width(), 1400/page.get_height()))
            try:
                with bitmap.to_pil() as image: image.save(folder/'page.png', format='PNG')
            finally:
                bitmap.close()
        finally:
            page.close()
        (folder/'page.json').write_text(json.dumps({'pages': len(book), 'page': index+1}), encoding='utf8')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
