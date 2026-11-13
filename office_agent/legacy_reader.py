from __future__ import annotations

import hashlib
import io
from pathlib import Path
import re
import struct
import sys
import tempfile

from .fidelity import _profile, _run, _temporary_root, runtime_path
from .rag_sources import MAX_SOURCE_BYTES, SourceError, _package


def validate_legacy(data, suffix):
    import olefile

    if suffix not in {'.doc', '.xls'} or not data.startswith(b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1') or len(data) > MAX_SOURCE_BYTES:
        raise SourceError('unsupported')
    try:
        with olefile.OleFileIO(io.BytesIO(data), raise_defects=olefile.DEFECT_INCORRECT) as archive:
            names = archive.listdir()
            common = {'\x05SummaryInformation', '\x05DocumentSummaryInformation', '\x01CompObj', '\x01Ole'}
            permitted = common | ({'WordDocument', '0Table', '1Table'} if suffix == '.doc' else {'Workbook', 'Book'})
            if not names or len(names) > 20 or any(len(name) != 1 or name[0] not in permitted for name in names):
                raise SourceError('unsupported')
            if sum(archive.get_size(name) for name in names) > MAX_SOURCE_BYTES:
                raise SourceError('capacity')
            streams = {}
            for name in names:
                if archive.get_size(name) > 16 * 1024 * 1024:
                    raise SourceError('capacity')
                streams[name[0]] = archive.openstream(name).read(16 * 1024 * 1024 + 1)
            if suffix == '.doc':
                word = streams.get('WordDocument', b'')
                if len(word) < 32 or word[:2] != b'\xec\xa5' or struct.unpack_from('<H', word, 10)[0] & 0x8100:
                    raise SourceError('unsupported')
                if ('1Table' if struct.unpack_from('<H', word, 10)[0] & 0x0200 else '0Table') not in streams:
                    raise SourceError('invalid')
                for value in streams.values():
                    normalized = value.replace(b'\x00', b'').upper()
                    if any(marker in normalized for marker in (b'DDE', b'INCLUDETEXT', b'INCLUDEPICTURE', b'HYPERLINK', b'JAVASCRIPT:', b'HTTP:', b'HTTPS:', b'FILE:', b'\\\\')):
                        raise SourceError('unsupported')
            else:
                workbook = streams.get('Workbook', streams.get('Book', b''))
                cursor = 0
                if not workbook or workbook[:2] not in {b'\x09\x08', b'\x09\x04', b'\x09\x02', b'\x09\x00'}:
                    raise SourceError('invalid')
                while cursor + 4 <= len(workbook):
                    tag, length = struct.unpack_from('<HH', workbook, cursor)
                    if tag == length == 0:
                        break
                    end = cursor + 4 + length
                    if end > len(workbook) or length > 8224:
                        raise SourceError('invalid')
                    if tag in {0x002F, 0x01AE, 0x0023, 0x00D3, 0x01BA, 0x0017, 0x005D, 0x0006, 0x0018, 0x04BC, 0x0221, 0x0236, 0x0036}:
                        raise SourceError('unsupported')
                    if tag == 0x0085 and (length < 6 or workbook[cursor + 9] != 0):
                        raise SourceError('unsupported')
                    if tag in {0x0809, 0x0409, 0x0209, 0x0009} and (length < 4 or struct.unpack_from('<H', workbook, cursor + 6)[0] not in {5, 16}):
                        raise SourceError('unsupported')
                    cursor = end
    except SourceError:
        raise
    except Exception:
        raise SourceError('invalid') from None


def read_legacy_document(path, *, include_tables=False, region=None, executable=None):
    from docx import Document
    from openpyxl import load_workbook
    from openpyxl.utils.cell import range_boundaries

    path = Path(path)
    if type(include_tables) is not bool or path.is_symlink() or path.stat().st_size > MAX_SOURCE_BYTES:
        raise SourceError('invalid')
    bounds = None
    if region is not None:
        if path.suffix.lower() != '.xls' or type(region) is not str or not re.fullmatch(r'[A-Za-z]{1,3}[1-9][0-9]{0,6}(?::[A-Za-z]{1,3}[1-9][0-9]{0,6})?', region):
            raise SourceError('invalid')
        bounds = range_boundaries(region)
        left, top, right, bottom = bounds
        if left > right or top > bottom or right > 16384 or bottom > 1048576 or right - left >= 20 or bottom - top >= 20:
            raise SourceError('capacity')
    with path.open('rb') as stream:
        data = stream.read(MAX_SOURCE_BYTES + 1)
    validate_legacy(data, path.suffix.lower())
    fingerprint = hashlib.sha256(data).hexdigest()
    processes = []
    target_suffix = '.docx' if path.suffix.lower() == '.doc' else '.xlsx'
    with tempfile.TemporaryDirectory(prefix='orl-', dir=_temporary_root()) as raw:
        scratch = Path(raw)
        copy = scratch / ('input' + path.suffix.lower())
        copy.write_bytes(data)
        destination = scratch / 'converted'
        destination.mkdir()
        _profile(destination / 'profile')
        _run((sys.executable, str(Path(__file__).with_name('fidelity_worker.py')), '--legacy',
              str(copy), str(destination), str(runtime_path(executable))), 30000, processes)
        converted = destination / ('input' + target_suffix)
        if not converted.is_file() or converted.stat().st_size > MAX_SOURCE_BYTES:
            raise SourceError('capacity')
        _package(converted.read_bytes(), target_suffix)
        truncated = False
        if target_suffix == '.docx':
            if region is not None:
                raise SourceError('invalid')
            document = Document(converted)
            paragraphs = [{'index': index + 1, 'text': paragraph.text[:4096], 'text_truncated': len(paragraph.text) > 4096}
                          for index, paragraph in enumerate(document.paragraphs[:20])]
            truncated = len(document.paragraphs) > 20 or any(item['text_truncated'] for item in paragraphs)
            result = {'type': 'word', 'paragraphs': paragraphs}
            if include_tables:
                tables, remaining = [], 400
                for index, table in enumerate(document.tables[:20]):
                    rows = []
                    for row_index, row in enumerate(table.rows[:20]):
                        cells = []
                        for column, cell in enumerate(row.cells[:min(20, remaining)]):
                            cells.append({'index': column + 1, 'text': cell.text[:4096], 'text_truncated': len(cell.text) > 4096})
                            remaining -= 1
                            truncated |= len(cell.text) > 4096
                        rows.append({'index': row_index + 1, 'cells': cells})
                        truncated |= len(row.cells) > len(cells)
                        if not remaining:
                            break
                    tables.append({'index': index + 1, 'rows': rows})
                    truncated |= len(table.rows) > len(rows)
                    if not remaining:
                        break
                truncated |= len(document.tables) > len(tables)
                result['tables'] = tables
        else:
            workbook = load_workbook(converted, read_only=True, data_only=False)
            try:
                sheets, remaining = [], 400
                for sheet in workbook.worksheets[:20]:
                    left, top, right, bottom = bounds or (1, 1, min(20, sheet.max_column or 1), min(20, sheet.max_row or 1))
                    cells = []
                    for row in sheet.iter_rows(min_row=top, max_row=bottom, min_col=left, max_col=right):
                        for cell in row:
                            if not remaining:
                                break
                            value = cell.value
                            if value is not None:
                                cells.append({'cell': cell.coordinate, 'value': str(value)[:4096] if isinstance(value, str) else value,
                                              'text_truncated': isinstance(value, str) and len(value) > 4096,
                                              'formula_result_unverified': cell.data_type == 'f'})
                            remaining -= 1
                        if not remaining:
                            break
                    sheets.append({'sheet': sheet.title, 'cells': cells})
                    truncated |= (sheet.max_row or 0) > bottom or (sheet.max_column or 0) > right
                    if not remaining:
                        break
                truncated |= len(sheets) < len(workbook.worksheets)
                result = {'type': 'excel', 'sheets': sheets, 'cached_values_verified': False}
            finally:
                workbook.close()
    with path.open('rb') as stream:
        fresh = stream.read(MAX_SOURCE_BYTES + 1)
    if hashlib.sha256(fresh).hexdigest() != fingerprint:
        raise SourceError('source_changed')
    return {**result, 'backend': 'libreoffice_read_only_conversion', 'source_sha256': fingerprint,
            'path': str(path), 'truncated': truncated, 'read_only': True, 'processes': processes,
            'limitations': ['Restricted passive legacy subset; macros, external field/link markers, names and BIFF formulas rejected',
                            'Converted representation only; original layout fidelity is unverified',
                            'No legacy write capability is enabled']}
