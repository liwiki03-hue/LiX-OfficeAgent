import json
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from office_agent.pdf_reader import read_pdf_payload
from office_agent.rag_sources import SourceError


def main():
    if len(sys.argv) != 5 or sys.argv[1] not in ('pdf', 'pdf-page', 'pdf-ocr'):
        return 2
    source, config, target = map(Path, sys.argv[2:])
    try:
        options = json.loads(config.read_text(encoding='utf-8'))
        if sys.argv[1] == 'pdf-ocr':
            from office_agent.research_ocr import ocr_payload
            report = ocr_payload(source, **options)
        elif sys.argv[1] == 'pdf-page':
            from office_agent.research_page import render_payload
            report = render_payload(source, **options)
        else:
            report = read_pdf_payload(source, **options)
    except SourceError as error:
        report = {'error': error.code}
    with target.open('x', encoding='utf-8') as stream:
        json.dump(report, stream, ensure_ascii=False)
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
