import json
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from office_agent.pdf_table import payload
from office_agent.rag_sources import SourceError


def main():
    folder = Path(sys.argv[1])
    try:
        options = json.loads((folder / 'options.json').read_text(encoding='utf8'))
        if 'proposal_box' in options:
            from office_agent.pdf_layout import propose
            result = propose(folder / 'input.pdf', options['page'], options['proposal_box'])
        else:
            result = payload(folder / 'input.pdf', **options)
    except SourceError as error:
        result = {'error': error.code}
    except Exception:
        result = {'error': 'unsupported'}
    with (folder / 'result.json').open('x', encoding='utf8') as stream:
        json.dump(result, stream, ensure_ascii=False)


if __name__ == '__main__':
    main()
