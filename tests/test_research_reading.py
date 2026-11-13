from io import BytesIO

import pypdfium2 as pdfium

from office_agent.research_reading import PdfCoverage
from office_agent.writes import digest
from test_fidelity import pdf_bytes
from test_grounded import setup_service
from test_archive_grounded import archive_source


def long_pdf():
    stream = BytesIO()
    with pdfium.PdfDocument.new() as writer:
        for index in range(7):
            text = f'Page {index + 1}. ' + ('The measured latency was 20 ms under laboratory conditions.' if index == 6 else 'Background only.')
            with pdfium.PdfDocument(pdf_bytes(text=text, width=600)) as source:
                writer.import_pages(source)
        writer.save(stream)
    return stream.getvalue()


def test_explicit_late_page_and_original_projection(tmp_path, monkeypatch):
    service, path, record, auth, busy = setup_service(tmp_path, monkeypatch)
    (tmp_path / 'paper.pdf').write_bytes(long_pdf())
    initial = service.capture(record, 'paper.pdf')['result']
    assert initial['coverage']['complete_pages'] == [1, 2, 3, 4, 5] and not initial['coverage']['complete']
    captured = service.capture(record, 'paper.pdf', 5, 2)['result']
    assert captured['coverage']['complete'] and captured['coverage']['complete_pages'] == list(range(1, 8))
    assert captured['capture_id'] != initial['capture_id']
    assert all(type(key) is str for key in captured['coverage']['intervals'])
    evidence = next(row for row in captured['fragments'] if row['page'] == 7)
    page = service.page(captured['capture_id'], evidence['id'])
    assert page['ok'], page
    assert page['result']['page'] == 7 and page['result']['image'].startswith('data:image/jpeg;base64,')
    assert page['result']['highlight'] and page['result']['precision'] == 'pdfium_page_to_device_run_bounds'
    assert page['result']['source_sha256'] == digest(tmp_path / 'paper.pdf')
    assert not service.page(captured['capture_id'], 'unowned-id')['ok']
    (tmp_path / 'paper.pdf').write_bytes(pdf_bytes(text='replacement'))
    assert service.page(captured['capture_id'], evidence['id'])['reason'] == 'professional_evidence_changed'


def test_archived_late_page_uses_receipt_and_changed_pdf_refuses(tmp_path, monkeypatch):
    service, path, record, auth, busy = setup_service(tmp_path, monkeypatch)
    saved = archive_source(service, record, long_pdf())
    capture = service.capture_archive(saved['save_id'], 6, 1)
    assert capture['ok'], capture
    capture = capture['result']
    assert capture['coverage']['read_pages'] == [7] and not capture['coverage']['complete']
    assert service.page(capture['capture_id'], capture['fragments'][0]['id'])['ok']
    (tmp_path / ('.office-rag-evidence-' + saved['save_id']) / 'fulltext.pdf').write_bytes(pdf_bytes(text='changed'))
    assert not service.page(capture['capture_id'], capture['fragments'][0]['id'])['ok']


def test_partial_intervals_never_invent_full_coverage_or_indexing():
    coverage = PdfCoverage()
    report = {'source_sha256': 'a' * 64, 'page_count': 2, 'next_page': 0, 'next_char_offset': 80,
              'pages': [{'page': 0, 'char_offset': 0, 'total_chars': 120, 'text_truncated': True, 'text_layer_empty': False,
                         'runs': [{'char_index': 0, 'char_count': 80}]}]}
    first = coverage.record('a', report)
    assert first['complete_pages'] == [] and first['index_coverage'] == 'not_published_by_reading'
    report['pages'][0].update(char_offset=80, text_truncated=False, runs=[{'char_index': 80, 'char_count': 40}])
    second = coverage.record('a', report)
    assert second['complete_pages'] == [1] and not second['complete'] and first['complete_pages'] == []
    report['source_sha256'] = 'b' * 64
    third = coverage.record('a', report)
    assert not third['complete_pages'] and third['intervals']['1'] == [[80, 120]]


def test_reading_failure_and_replacement_are_visible(tmp_path, monkeypatch):
    service, path, record, auth, busy = setup_service(tmp_path, monkeypatch)
    assert service.capture(record, 'paper.pdf')['ok']

    def failed(*args, **kwargs):
        raise RuntimeError('parser unavailable')

    monkeypatch.setattr('office_agent.pdf_reader.read_pdf', failed)
    assert not service.capture(record, 'paper.pdf')['ok']
    status = service.reading_status()['result']['sources'][0]
    assert status['failed_pages'] == [1] and status['failed_ranges'][0]['start_page'] == 1
    assert not status['stale']
    (tmp_path / 'paper.pdf').write_bytes(pdf_bytes(text='replacement'))
    status = service.reading_status()['result']['sources'][0]
    assert status['stale'] and not status['complete_current_version']
