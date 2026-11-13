import csv
import hashlib
import json

import pytest

from tools.review_table_context import CHECKS, evaluate


def packet(tmp_path, complete=False, origin='independent_human'):
    folder = tmp_path / 'blind-review'
    folder.mkdir()
    (folder / 'TASKS.json').write_text(json.dumps([{'case_id': 'R01'}]), encoding='utf8')
    hashes = {'TASKS.json': hashlib.sha256((folder / 'TASKS.json').read_bytes()).hexdigest()}
    (tmp_path / 'ACCEPTANCE.json').write_text(json.dumps({'packet_sha256': hashes}), encoding='utf8')
    declaration = {'reviewer_id': 'synthetic-contract-fixture', 'is_independent': True if complete else None,
        'participated_in_development': False if complete else None, 'reviewed_on': '2026-10-19' if complete else '',
        'checks': {key: True if complete else None for key in CHECKS}}
    (folder / 'REVIEWER.json').write_text(json.dumps(declaration), encoding='utf8')
    with (folder / 'annotations.csv').open('w', encoding='utf8', newline='') as stream:
        writer = csv.writer(stream)
        writer.writerow(['case_id', 'reviewer', 'label_origin', 'correct', 'false_positive', 'false_negative', 'correction_count', 'seconds', 'notes'])
        writer.writerow(['R01', 'synthetic-contract-fixture', origin, '1', '0', '0', '0', '1', 'unit fixture only'] if complete else ['R01', *([''] * 8)])
    return folder


def test_empty_human_packet_is_pending_with_no_accuracy(tmp_path):
    packet(tmp_path)
    report = evaluate(tmp_path)
    assert report['status'] == 'pending' and report['accuracy'] is None and report['independent_cases_completed'] == 0


@pytest.mark.parametrize('origin', ['developer', 'model'])
def test_other_labels_cannot_be_imported_as_independent(tmp_path, origin):
    packet(tmp_path, True, origin)
    with pytest.raises(ValueError, match='Independent human'):
        evaluate(tmp_path)


def test_changed_reviewed_output_refused(tmp_path):
    folder = packet(tmp_path, True)
    (folder / 'TASKS.json').write_text('[]', encoding='utf8')
    with pytest.raises(ValueError, match='packet changed'):
        evaluate(tmp_path)


def test_manual_failure_cannot_pass_on_semantic_labels(tmp_path):
    folder = packet(tmp_path, True)
    declaration = json.loads((folder / 'REVIEWER.json').read_text(encoding='utf8'))
    declaration['checks']['word_comments_visible'] = False
    (folder / 'REVIEWER.json').write_text(json.dumps(declaration), encoding='utf8')
    assert evaluate(tmp_path)['status'] == 'failed'
