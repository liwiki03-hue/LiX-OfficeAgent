import json
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from office_agent.webui.document_open import parse_payload


def main():
    if len(sys.argv) != 3:
        return 2
    source, target = map(Path, sys.argv[1:])
    try:
        report = parse_payload(source)
    except Exception:
        report = {'error': 'parse_failed'}
    data = json.dumps(report, ensure_ascii=False).encode('utf-8')
    if len(data) > 1024 * 1024:
        data = b'{"error":"view_capacity"}'
    with target.open('xb') as stream:
        stream.write(data)
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
