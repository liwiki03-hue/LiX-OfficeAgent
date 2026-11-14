import argparse
import json
from pathlib import Path
import sys

if not __package__:
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from office_agent.download_cache import DownloadCache


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--root', type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument('--kind', choices=['frontend', 'renderer'], required=True)
    parser.add_argument('--confirm-quarantine', action='store_true')
    args = parser.parse_args()
    cache = DownloadCache(args.root)
    plan = cache.preview(args.kind)
    result = cache.quarantine(plan['id'], True) if args.confirm_quarantine and plan['repairable'] else plan
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
