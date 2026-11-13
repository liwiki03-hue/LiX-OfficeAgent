"""Host-only, read-only RAG sources. No scanning, providers or persistence.

Names are canonical root-relative filenames; doc identifiers are in-memory bindings,
never paths. Checks detect ordinary changes, not an OS isolation boundary against
same-privilege concurrent filesystem attackers.
"""
from __future__ import annotations

import hashlib
import io
import os
import posixpath
import re
import stat
import zipfile
from collections import OrderedDict
from copy import deepcopy
from pathlib import Path, PurePosixPath

from lxml import etree

from .rag import MAX_BLOCKS, MAX_TEXT, MAX_TOTAL_TEXT, RagError, structure_chunks
from .writes import Workspace

MAX_SOURCE_BYTES = 32 * 1024 * 1024
MAX_ZIP_ENTRIES = 2048
MAX_ENTRY_BYTES = 16 * 1024 * 1024
MAX_ZIP_BYTES = 64 * 1024 * 1024
MAX_COMPRESSION_RATIO = 100
MAX_XML_NODES = 200_000
MAX_IDENTITIES = 256
PARSER_VERSION = 'native-structure-v1'
CACHE_CHARACTERS = 2_000_000
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
S = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
P = '{http://schemas.openxmlformats.org/package/2006/relationships}'
C = '{http://schemas.openxmlformats.org/package/2006/content-types}'


class SourceError(ValueError):
    """Only fixed redacted codes cross the host boundary."""
    def __init__(self, code='invalid'):
        self.code = code if code in {'source_changed', 'invalid', 'capacity', 'unsupported', 'sensitive', 'pdf_text_truncated', 'pdf_text_missing'} else 'invalid'
        super().__init__(self.code)


def _identity(info):
    return (info.st_dev, info.st_ino, info.st_size, info.st_mtime_ns, info.st_ctime_ns)


def _ordinary(info, directory=False):
    if (info.st_file_attributes if hasattr(info, 'st_file_attributes') else 0) & 0x400:
        raise SourceError('invalid')
    if not (stat.S_ISDIR(info.st_mode) if directory else stat.S_ISREG(info.st_mode)):
        raise SourceError('invalid')
    if not directory and info.st_nlink != 1:
        raise SourceError('invalid')


def _xml(data):
    parser = etree.XMLParser(resolve_entities=False, load_dtd=False, no_network=True,
                             huge_tree=False, recover=False)
    root = etree.fromstring(data, parser)
    if root.getroottree().docinfo.doctype or any(isinstance(n, etree._Entity) for n in root.iter()):
        raise SourceError('unsupported')
    count = 0
    for node in root.iter():
        count += 1
        if count > MAX_XML_NODES:
            raise SourceError('capacity')
        if not isinstance(node.tag, str):
            continue
        local = etree.QName(node).localname.lower()
        if local in {'instrtext', 'fldsimple'}:
            instruction = ''.join(node.itertext()) + ' ' + ' '.join(node.attrib.values())
            if re.search(r'(?i)\b(?:DDE|DDEAUTO|INCLUDETEXT|INCLUDEPICTURE|LINK|HYPERLINK)\b', instruction):
                raise SourceError('unsupported')
        if local in {'externalreference', 'externalreferences', 'externalbook', 'altchunk'}:
            raise SourceError('unsupported')
        if any(etree.QName(k).localname == 'href' for k in node.attrib):
            raise SourceError('unsupported')
    return root


def _package(data, suffix):
    """Validate every member, including unused XML parts, before extracting text."""
    parts, xml_nodes = {}, 0
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        entries = archive.infolist()
        if len(entries) > MAX_ZIP_ENTRIES:
            raise SourceError('capacity')
        total = 0
        for entry in entries:
            name = entry.filename
            path = PurePosixPath(name)
            if (not name or len(name) > 256 or '\\' in name or ':' in name or
                    path.is_absolute() or '..' in path.parts or name != str(path) or
                    name in parts or stat.S_ISLNK(entry.external_attr >> 16)):
                raise SourceError('invalid')
            if entry.flag_bits & 1 or entry.compress_type not in {zipfile.ZIP_STORED, zipfile.ZIP_DEFLATED}:
                raise SourceError('unsupported')
            total += entry.file_size
            if (entry.file_size > MAX_ENTRY_BYTES or total > MAX_ZIP_BYTES or
                    entry.file_size > MAX_COMPRESSION_RATIO * max(1, entry.compress_size)):
                raise SourceError('capacity')
            if any(x in name.lower() for x in ('vbaproject', 'externalLinks'.lower(), 'macros', 'encryptedpackage')):
                raise SourceError('unsupported')
            with archive.open(entry) as stream:
                value = stream.read(MAX_ENTRY_BYTES + 1)
            if len(value) != entry.file_size:
                raise SourceError('invalid')
            parts[name] = _xml(value) if name.endswith(('.xml', '.rels')) else value
            if isinstance(parts[name], etree._Element):
                xml_nodes += sum(1 for _ in parts[name].iter())
                if xml_nodes > MAX_XML_NODES:
                    raise SourceError('capacity')
    types = parts.get('[Content_Types].xml')
    if not isinstance(types, etree._Element) or types.tag != C + 'Types':
        raise SourceError('invalid')
    formats = {
        '.docx': ('word/document.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'),
        '.xlsx': ('xl/workbook.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'),
        '.pptx': ('ppt/presentation.xml', 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml'),
    }
    if suffix not in formats:
        raise SourceError('unsupported')
    core, expected = formats[suffix]
    if not any(n.get('PartName') == '/' + core and n.get('ContentType') == expected for n in types):
        raise SourceError('invalid')
    for node in types:
        if 'macro' in node.get('ContentType', '').lower():
            raise SourceError('unsupported')
    for name, root in parts.items():
        if name.endswith('.rels'):
            if root.tag != P + 'Relationships':
                raise SourceError('invalid')
            ids = set()
            for rel in root:
                identifier = rel.get('Id', '')
                if not identifier or identifier in ids:
                    raise SourceError('invalid')
                ids.add(identifier)
                target = rel.get('Target', '')
                if (rel.tag != P + 'Relationship' or rel.get('TargetMode', 'Internal') != 'Internal' or
                        not target or ':' in target or '\\' in target or target.startswith('//')):
                    raise SourceError('unsupported')
                base = posixpath.dirname(posixpath.dirname(name)) if name != '_rels/.rels' else ''
                resolved = posixpath.normpath(posixpath.join(base, target.lstrip('/'))) if not target.startswith('/') else target.lstrip('/')
                if resolved.startswith('../') or resolved not in parts:
                    raise SourceError('invalid')
    return parts


class SourceExtractor:
    def __init__(self, workspace):
        if type(workspace) is not Workspace:
            raise SourceError('invalid')
        self._workspace = workspace
        self._root = workspace.root
        try:
            _ordinary(self._root.lstat(), True)
            info = self._root.stat()
            self._root_id = (info.st_dev, info.st_ino)
        except SourceError:
            raise
        except Exception:
            raise SourceError('invalid') from None
        self._bindings = {}
        self._cache = OrderedDict()
        self._cache_characters = 0

    def clear_cache(self):
        self._cache.clear()
        self._cache_characters = 0

    def _path(self, name):
        if (self._workspace.root != self._root or self._root.resolve(strict=True) != self._root):
            raise SourceError('source_changed')
        root_info = self._root.lstat()
        _ordinary(root_info, True)
        if (root_info.st_dev, root_info.st_ino) != self._root_id:
            raise SourceError('source_changed')
        path = self._root / name
        for ancestor in reversed(path.parents):
            if ancestor == self._root or ancestor.is_relative_to(self._root):
                _ordinary(ancestor.lstat(), True)
        _ordinary(path.lstat())
        guarded = self._workspace.file(str(path))
        if guarded != path:
            raise SourceError('invalid')
        return path

    def _snapshot(self, name):
        path = self._path(name)
        before = path.lstat()
        if before.st_size > MAX_SOURCE_BYTES:
            raise SourceError('capacity')
        with path.open('rb') as stream:
            opened = os.fstat(stream.fileno())
            _ordinary(opened)
            # Windows path/handle ctime semantics differ; keep full named stamps.
            if _identity(opened)[:4] != _identity(before)[:4]:
                raise SourceError('source_changed')
            data = stream.read(MAX_SOURCE_BYTES + 1)
            after = os.fstat(stream.fileno())
        if len(data) > MAX_SOURCE_BYTES:
            raise SourceError('capacity')
        fresh = self._path(name).lstat()
        _ordinary(after)
        if _identity(before)[:4] != _identity(after)[:4] or _identity(before) != _identity(fresh) or len(data) != before.st_size:
            raise SourceError('source_changed')
        return data, _identity(before), hashlib.sha256(data).hexdigest()

    def extract(self, name):
        try:
            if (type(name) is not str or not name or len(name) > 128 or '\\' in name or ':' in name or
                    '\x00' in name or PurePosixPath(name).is_absolute() or
                    any(p in {'', '.', '..'} or p != p.rstrip(' .') for p in name.split('/'))):
                raise SourceError('invalid')
            if any(p.lower().startswith('.env') for p in name.split('/')):
                raise SourceError('sensitive')
            suffix = PurePosixPath(name).suffix.lower()
            if suffix not in {'.txt', '.docx', '.xlsx'}:
                raise SourceError('unsupported')
            doc_id = hashlib.sha256(name.casefold().encode('utf-8')).hexdigest()
            if doc_id not in self._bindings and len(self._bindings) >= MAX_IDENTITIES:
                raise SourceError('capacity')
            data, identity, fingerprint = self._snapshot(name)
            key = (name, identity, fingerprint, PARSER_VERSION)
            cached = self._cache.get(key)
            blocks = deepcopy(cached) if cached is not None else self._parse(data, suffix, name)
            if not blocks:
                raise SourceError('invalid')
            structure_chunks(blocks)
            _, after_identity, after_fingerprint = self._snapshot(name)
            if after_identity != identity or after_fingerprint != fingerprint:
                raise SourceError('source_changed')
            if cached is not None:
                self._cache.move_to_end(key)
            else:
                for old in list(self._cache):
                    if old[0] == name:
                        self._cache_characters -= sum(len(block['text']) for block in self._cache.pop(old))
                size = sum(len(block['text']) for block in blocks)
                while self._cache and (len(self._cache) >= 20 or self._cache_characters + size > CACHE_CHARACTERS):
                    _, removed = self._cache.popitem(last=False)
                    self._cache_characters -= sum(len(block['text']) for block in removed)
                if size <= CACHE_CHARACTERS:
                    self._cache[key] = deepcopy(blocks)
                    self._cache_characters += size
            self._bindings[doc_id] = (name, identity)
            return {'doc_id': doc_id, 'fingerprint': fingerprint, 'blocks': blocks}
        except SourceError:
            raise
        except RagError as error:
            code = 'sensitive' if 'sensitive' in error.code else 'capacity' if error.code == 'rag_capacity' else 'invalid'
            raise SourceError(code) from None
        except Exception:
            raise SourceError('invalid') from None

    def fingerprint(self, doc_id):
        try:
            if type(doc_id) is not str or doc_id not in self._bindings:
                raise SourceError('invalid')
            name, original = self._bindings[doc_id]
            _, identity, fingerprint = self._snapshot(name)
            if identity[:2] != original[:2]:
                raise SourceError('source_changed')
            return fingerprint
        except SourceError:
            raise
        except Exception:
            raise SourceError('source_changed') from None

    def _parse(self, data, suffix, source):
        blocks, total = [], 0
        def add(kind, text, location, group):
            nonlocal total
            if not text.strip():
                return
            total += len(text)
            if len(text) > MAX_TEXT or total > MAX_TOTAL_TEXT or len(blocks) >= MAX_BLOCKS:
                raise SourceError('capacity')
            if len(location) > 128 or len(group) > 64:
                raise SourceError('capacity')
            blocks.append(dict(source=source, location=location, type=kind, text=text, group=group))
        if suffix == '.txt':
            for index, line in enumerate(data.decode('utf-8-sig').splitlines(), 1):
                add('paragraph', line, f'line:{index}', 'text')
            return blocks
        parts = _package(data, suffix)
        if suffix == '.docx':
            root = parts.get('word/document.xml')
            if root is None or root.tag != W + 'document':
                raise SourceError('invalid')
            body = root.find(W + 'body')
            if body is None:
                raise SourceError('invalid')
            def text(node):
                return ''.join((n.text or '') if n.tag == W + 't' else '\t' if n.tag == W + 'tab' else '\n'
                               for n in node.iter() if n.tag in {W + 't', W + 'tab', W + 'br', W + 'cr'})
            for position, node in enumerate(body, 1):
                if node.tag == W + 'p':
                    style = node.find(W + 'pPr/' + W + 'pStyle')
                    outline = node.find(W + 'pPr/' + W + 'outlineLvl')
                    heading = outline is not None or (style is not None and re.fullmatch(r'Heading\d+', style.get(W + 'val', ''), re.I))
                    add('heading' if heading else 'paragraph', text(node), f'body:{position}/paragraph', 'body')
                elif node.tag == W + 'tbl':
                    if any(n.tag == W + 'tbl' for n in list(node.iter())[1:]):
                        raise SourceError('unsupported')
                    for row_index, row in enumerate(node.findall(W + 'tr'), 1):
                        cells = ['\n'.join(text(p) for p in cell.findall(W + 'p')) for cell in row.findall(W + 'tc')]
                        if any(value.strip() for value in cells):
                            add('table_row', ' | '.join(cells),
                                f'body:{position}/table/row:{row_index}', f'table:{position}')
                elif node.tag != W + 'sectPr':
                    raise SourceError('unsupported')
            return blocks
        root = parts.get('xl/workbook.xml')
        rels = parts.get('xl/_rels/workbook.xml.rels')
        if root is None or root.tag != S + 'workbook' or rels is None:
            raise SourceError('invalid')
        targets = {n.get('Id'): n for n in rels}
        shared = parts.get('xl/sharedStrings.xml')
        strings = []
        if shared is not None:
            if shared.tag != S + 'sst':
                raise SourceError('invalid')
            for item in shared:
                value = ''.join(n.text or '' for n in item.iter(S + 't'))
                if len(value) > MAX_TEXT or len(strings) >= MAX_XML_NODES:
                    raise SourceError('capacity')
                strings.append(value)
        sheets = root.find(S + 'sheets')
        if sheets is None:
            raise SourceError('invalid')
        for index, sheet in enumerate(sheets, 1):
            title = sheet.get('name', '')
            if not title or len(title) > 31:
                raise SourceError('invalid')
            rel = targets.get(sheet.get(R + 'id'))
            if rel is None or rel.get('Type') != R[1:-1] + '/worksheet':
                raise SourceError('invalid')
            target = rel.get('Target', '')
            part = target.lstrip('/') if target.startswith('/') else posixpath.normpath('xl/' + target)
            worksheet = parts.get(part)
            if worksheet is None or worksheet.tag != S + 'worksheet':
                raise SourceError('invalid')
            rows = worksheet.find(S + 'sheetData')
            if rows is None:
                raise SourceError('invalid')
            previous = 0
            cell_count = 0
            for row in rows:
                number = int(row.get('r', '0'))
                if row.tag != S + 'row' or not previous < number <= 1048576:
                    raise SourceError('invalid')
                previous = number
                values, seen = [], set()
                for cell in row:
                    cell_count += 1
                    if cell_count > MAX_XML_NODES:
                        raise SourceError('capacity')
                    coordinate = cell.get('r', '')
                    match = re.fullmatch(r'([A-Z]{1,3})([1-9][0-9]{0,6})', coordinate)
                    if cell.tag != S + 'c' or not match or int(match[2]) != number or coordinate in seen:
                        raise SourceError('invalid')
                    column = 0
                    for letter in match[1]:
                        column = column * 26 + ord(letter) - 64
                    if column > 16384:
                        raise SourceError('invalid')
                    seen.add(coordinate)
                    formula, value = cell.find(S + 'f'), cell.find(S + 'v')
                    if formula is not None:
                        expression = formula.text or ''
                        if not expression or '[' in expression or '|' in expression or re.search(r'(?i)\b(?:DDE|WEBSERVICE|HYPERLINK|RTD)\s*\(', expression):
                            raise SourceError('unsupported')
                        rendered = '=' + expression
                    elif cell.get('t') == 'inlineStr':
                        rendered = ''.join(n.text or '' for n in cell.iter(S + 't'))
                    elif cell.get('t') == 's':
                        rendered = strings[int(value.text)] if value is not None and value.text is not None and int(value.text) >= 0 else None
                        if rendered is None:
                            raise SourceError('invalid')
                    else:
                        rendered = value.text or '' if value is not None else ''
                    if rendered.strip():
                        values.append(f'{coordinate}: {rendered}')
                add('sheet_row', ' | '.join(values), f'sheet:{index}:{title}/row:{number}', f'sheet:{index}:{title}')
        return blocks
