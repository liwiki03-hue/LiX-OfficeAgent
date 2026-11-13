import csv
import hashlib
import json
import math
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
PACKET = ROOT / 'deliverables/table-context-20261019'
CHECKS = {'word_opens', 'word_comments_visible', 'word_original_text_preserved', 'word_non_target_preserved',
          'word_save_reopen', 'webview2_manual_workflow', 'desktop_restart_manual_workflow'}


def evaluate(folder):
    acceptance = json.loads((folder / 'ACCEPTANCE.json').read_text(encoding='utf8'))
    packet = folder / 'blind-review'
    for name, expected in acceptance['packet_sha256'].items():
        if name not in {'review-source.docx', 'review-commented.docx', 'APPLICATION_OUTPUTS.json', 'TASKS.json'}:
            raise ValueError('Invalid packet manifest')
        if hashlib.sha256((packet / name).read_bytes()).hexdigest() != expected:
            raise ValueError('Review packet changed; repeat validation before reviewing')
    for name, expected in acceptance.get('material_sha256', {}).items():
        if name not in {'NIST.SP.260-230.pdf', 'table-50-transcribed.xlsx', 'SOURCE.json'}:
            raise ValueError('Invalid source manifest')
        if hashlib.sha256((folder / 'materials' / name).read_bytes()).hexdigest() != expected:
            raise ValueError('Review source changed; repeat validation before reviewing')
    reviewer = json.loads((packet / 'REVIEWER.json').read_text(encoding='utf8'))
    if set(reviewer) != {'reviewer_id', 'is_independent', 'participated_in_development', 'reviewed_on', 'checks'} or set(reviewer['checks']) != CHECKS:
        raise ValueError('Invalid reviewer declaration')
    tasks = json.loads((packet / 'TASKS.json').read_text(encoding='utf8'))
    with (packet / 'annotations.csv').open(encoding='utf-8-sig', newline='') as stream:
        rows = list(csv.DictReader(stream))
    if len(rows) != len(tasks) or {row['case_id'] for row in rows} != {row['case_id'] for row in tasks}:
        raise ValueError('Incomplete or duplicated task IDs')
    complete = []
    for row in rows:
        required = {'case_id', 'reviewer', 'label_origin', 'correct', 'false_positive', 'false_negative', 'correction_count', 'seconds', 'notes'}
        if set(row) != required:
            raise ValueError('Invalid annotation fields')
        if not all(row[key].strip() for key in required - {'notes'}):
            continue
        if (reviewer['is_independent'] is not True or reviewer['participated_in_development'] is not False or
                not reviewer['reviewer_id'] or row['reviewer'] != reviewer['reviewer_id'] or
                row['label_origin'] != 'independent_human' or not reviewer['reviewed_on']):
            raise ValueError('Independent human declaration required; developer/model labels are not interchangeable')
        if any(row[key] not in {'0', '1'} for key in ('correct', 'false_positive', 'false_negative')):
            raise ValueError('Invalid review labels')
        if int(row['correction_count']) < 0 or not math.isfinite(float(row['seconds'])) or float(row['seconds']) < 0:
            raise ValueError('Invalid review measurement')
        complete.append(row)
    if any(value is not None and type(value) is not bool for value in reviewer['checks'].values()):
        raise ValueError('Manual checks must be true, false or null')
    failures = sum(row['correct'] != '1' or row['false_positive'] != '0' or row['false_negative'] != '0' for row in complete)
    manual_failed = any(value is False for value in reviewer['checks'].values())
    finished = len(complete) == len(rows) and all(value is not None for value in reviewer['checks'].values())
    status = 'failed' if failures or manual_failed else 'passed' if finished else 'pending'
    return {'status': status, 'independent_cases_completed': len(complete), 'pending_cases': len(rows) - len(complete),
            'accuracy': sum(int(row['correct']) for row in complete) / len(complete) if complete else None,
            'false_positive': sum(int(row['false_positive']) for row in complete) if complete else None,
            'false_negative': sum(int(row['false_negative']) for row in complete) if complete else None,
            'manual_checks_completed': sum(type(value) is bool for value in reviewer['checks'].values()),
            'manual_checks_required': len(CHECKS), 'packet_hashes_match': True, 'identity_basis': 'reviewer_declaration_not_external_identity_verification'}


if __name__ == '__main__':
    print(json.dumps(evaluate(PACKET), ensure_ascii=False, indent=2))
