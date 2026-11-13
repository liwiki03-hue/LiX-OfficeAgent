from __future__ import annotations

import ast
from copy import deepcopy
import csv
from io import BytesIO, StringIO
import json
import math
import os
from pathlib import Path
import re
import shutil
import sys
import tempfile
from typing import Literal
import uuid
from lxml import etree as ET
import zipfile

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches
from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries, get_column_letter
from PIL import Image
from pydantic import Field

from .writes import GuardError, Plan, Strict, digest
from .professional_edits import ConstrainedDocx
from .range_edits import RangeDocx
from .structured_edits import StructuredDocx


class WordNumberedList(Strict):
    kind: Literal['word_numbered_list'] = 'word_numbered_list'
    path: str
    items: list[str] = Field(min_length=1, max_length=20)


class WordLocalImage(Strict):
    kind: Literal['word_local_image'] = 'word_local_image'
    path: str
    image_name: str = Field(min_length=1, max_length=200)
    width_inches: float = Field(default=3.0, ge=0.2, le=6.0)


class WordToc(Strict):
    kind: Literal['word_toc'] = 'word_toc'
    path: str


class PptxReplace(Strict):
    kind: Literal['pptx_replace'] = 'pptx_replace'
    path: str
    slide_index: int = Field(ge=0, le=99)
    shape_id: int = Field(ge=1)
    old: str = Field(min_length=1, max_length=2000)
    new: str = Field(max_length=2000)


class PptxSlide(Strict):
    kind: Literal['pptx_slide'] = 'pptx_slide'
    path: str
    action: Literal['add_blank', 'delete']
    slide_index: int = Field(default=0, ge=0, le=99)


class ExcelRecalculate(Strict):
    kind: Literal['excel_recalculate'] = 'excel_recalculate'
    path: str


class DocumentExport(Strict):
    kind: Literal['document_export'] = 'document_export'
    path: str
    format: Literal['pdf', 'csv']
    sheet: str | None = Field(default=None, max_length=31)


REQUESTS = (WordNumberedList, WordLocalImage, WordToc, PptxReplace, PptxSlide,
            ExcelRecalculate, DocumentExport, ConstrainedDocx, StructuredDocx)
TOOL_MODELS = {'plan_numbered_list': WordNumberedList, 'plan_local_image': WordLocalImage,
               'plan_toc': WordToc, 'plan_pptx_replace': PptxReplace,
               'plan_pptx_slide': PptxSlide, 'plan_recalculate': ExcelRecalculate,
               'plan_export_copy': DocumentExport}
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
S = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
REL = '{http://schemas.openxmlformats.org/package/2006/relationships}'
CT = '{http://schemas.openxmlformats.org/package/2006/content-types}'
TOC = 'TOC \\o "1-3" \\h \\z \\u'


def _parts(path, *, range_edit=False):
    from .fidelity import package_projection, FidelityError
    try:
        package_projection(path, for_render=not range_edit)
    except FidelityError as error:
        raise GuardError(error.code) from error
    with zipfile.ZipFile(path) as archive:
        result = {name: archive.read(name) for name in archive.namelist()}
    if range_edit:
        from .range_edits import static_render_parts
        static_render_parts(result)
    return result


def _zip(parts, target):
    with zipfile.ZipFile(target, 'x', zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)


def _xml(node):
    return ET.tostring(node, encoding='utf-8', xml_declaration=True)


def _image(space, source, request):
    if Path(request.image_name).name != request.image_name or '/' in request.image_name or '\\' in request.image_name:
        raise GuardError('image must be a same-directory file name')
    path = space.file(str(source.parent / request.image_name))
    if path.parent != source.parent or path.suffix.lower() not in {'.png', '.jpg', '.jpeg'} or path.stat().st_size > 5_000_000:
        raise GuardError('bounded local PNG/JPEG required')
    with Image.open(path) as image:
        if image.format not in {'PNG', 'JPEG'} or image.width * image.height > 8_000_000 or getattr(image, 'n_frames', 1) != 1:
            raise GuardError('simple bounded image required')
        image.verify()
    return path


def _word(parts):
    document = ET.fromstring(parts['word/document.xml'])
    forbidden = {'ins', 'del', 'sdt', 'altChunk', 'object', 'drawing', 'pict'}
    if any(node.tag.rsplit('}', 1)[-1] in forbidden for node in document.iter()):
        raise GuardError('advanced Word structures require a high-fidelity adapter')
    return document


def _slides(path, parts):
    from .presentations import _load
    _, slides, _ = _load(path)
    if not 1 <= len(slides) <= 100:
        raise GuardError('simple deck of 1-100 slides required')
    presentation = ET.fromstring(parts['ppt/presentation.xml'])
    if any(node.tag.rsplit('}', 1)[-1] in {'sectionLst', 'custShowLst', 'extLst', 'custDataLst'} for node in presentation.iter()):
        raise GuardError('presentation sections, custom shows and extensions are read-only')
    for name in slides:
        root = ET.fromstring(parts[name])
        if any(node.tag in {P+'graphicFrame', P+'pic', P+'grpSp', P+'cxnSp', P+'timing', P+'transition', A+'fld'} for node in root.iter()):
            raise GuardError('complex slides are read-only')
        relation_name = name.rsplit('/', 1)[0] + '/_rels/' + name.rsplit('/', 1)[1] + '.rels'
        if relation_name in parts:
            relations = ET.fromstring(parts[relation_name])
            if any(not node.get('Type', '').endswith('/slideLayout') for node in relations):
                raise GuardError('slide notes, media or links are read-only')
    return slides


def _numeric(value):
    if type(value) not in {int, float} or not math.isfinite(value) or abs(value) > 1e100:
        raise GuardError('only finite bounded numeric formula values supported')
    return value


def formula_results(path):
    workbook = load_workbook(path, data_only=False)
    if len(workbook.worksheets) > 20:
        raise GuardError('too many worksheets')
    results = {}
    count = 0
    for sheet in workbook:
        if sheet.max_row * sheet.max_column > 10000:
            raise GuardError('recalculation limited to 10000 cells per sheet')
        cache, active = {}, set()
        def cell_value(address):
            address = address.replace('$', '')
            if address in cache:
                return cache[address]
            if address in active or len(active) >= 32:
                raise GuardError('cyclic or excessively deep formula')
            active.add(address)
            value = sheet[address].value
            if isinstance(value, str) and value.startswith('='):
                if len(value) > 500 or any(token in value for token in ('!', '[', ']', '|', '"', "'", ';')):
                    raise GuardError('unsupported formula syntax')
                text = value[1:].upper()
                ranges = {}
                def range_token(match):
                    bounds = range_boundaries(match.group().replace('$', ''))
                    left, top, right, bottom = bounds
                    if (right-left+1)*(bottom-top+1) > 1000:
                        raise GuardError('formula range too large')
                    token = f'RANGE_{len(ranges)}'
                    ranges[token] = [cell_value(f'{get_column_letter(col)}{row}') for row in range(top, bottom+1) for col in range(left, right+1)]
                    return token
                text = re.sub(r'\$?[A-Z]{1,3}\$?[1-9][0-9]*:\$?[A-Z]{1,3}\$?[1-9][0-9]*', range_token, text)
                text = re.sub(r'(?<![A-Z_0-9])\$?([A-Z]{1,3})\$?([1-9][0-9]*)(?![A-Z_0-9(])', lambda m: f'CELL_{m[1]}{m[2]}', text)
                try:
                    tree = ast.parse(text, mode='eval')
                except SyntaxError:
                    raise GuardError('unsupported formula syntax') from None
                if sum(1 for _ in ast.walk(tree)) > 150:
                    raise GuardError('formula too complex')
                def evaluate(node):
                    if isinstance(node, ast.Expression):
                        return evaluate(node.body)
                    if isinstance(node, ast.Constant):
                        return _numeric(node.value)
                    if isinstance(node, ast.Name):
                        if node.id in ranges:
                            return ranges[node.id]
                        if node.id.startswith('CELL_'):
                            return cell_value(node.id[5:])
                    if isinstance(node, ast.UnaryOp) and isinstance(node.op, (ast.UAdd, ast.USub)):
                        return _numeric(evaluate(node.operand) * (-1 if isinstance(node.op, ast.USub) else 1))
                    if isinstance(node, ast.BinOp):
                        left, right = _numeric(evaluate(node.left)), _numeric(evaluate(node.right))
                        operations = {ast.Add: lambda: left+right, ast.Sub: lambda: left-right, ast.Mult: lambda: left*right, ast.Div: lambda: left/right}
                        operation = operations.get(type(node.op))
                        if operation:
                            return _numeric(operation())
                    if isinstance(node, ast.Call) and isinstance(node.func, ast.Name) and not node.keywords:
                        args = [evaluate(arg) for arg in node.args]
                        values = [value for arg in args for value in (arg if isinstance(arg, list) else [arg])]
                        functions = {'SUM': lambda: sum(values), 'AVERAGE': lambda: sum(values)/len(values), 'MIN': lambda: min(values), 'MAX': lambda: max(values), 'COUNT': lambda: len(values)}
                        if node.func.id in functions and values:
                            return _numeric(functions[node.func.id]())
                        if node.func.id == 'ABS' and len(args) == 1 and not isinstance(args[0], list):
                            return abs(_numeric(args[0]))
                    raise GuardError('formula outside verified arithmetic/SUM/AVERAGE/MIN/MAX/COUNT/ABS subset')
                try:
                    value = _numeric(evaluate(tree))
                except (ZeroDivisionError, OverflowError, ValueError) as exc:
                    raise GuardError('formula cannot be independently verified') from exc
            elif value is None:
                value = 0
            else:
                value = _numeric(value)
            active.remove(address)
            cache[address] = value
            return value
        for row in sheet:
            for cell in row:
                if cell.data_type == 'f':
                    count += 1
                    if count > 100:
                        raise GuardError('at most 100 formulas per recalculation')
                    results.setdefault(sheet.title, {})[cell.coordinate] = cell_value(cell.coordinate)
    workbook.close()
    if not results:
        raise GuardError('no supported formulas to recalculate')
    return results


def _csv_value(value):
    if isinstance(value, str) and value.lstrip().startswith(('=', '+', '-', '@', '\t', '\r')):
        raise GuardError('CSV formula injection risk; export refused')
    if value is not None and type(value) not in {str, int, float, bool}:
        raise GuardError('CSV supports ordinary literal values only')
    return '' if value is None else str(value)


def csv_text(path, sheet_name):
    workbook = load_workbook(path, data_only=False)
    if sheet_name not in workbook.sheetnames:
        raise GuardError('explicit existing worksheet required for CSV export')
    sheet = workbook[sheet_name]
    if sheet.max_row * sheet.max_column > 10000 or sheet.merged_cells:
        raise GuardError('CSV export limited to 10000 unmerged cells')
    target = StringIO(newline='')
    writer = csv.writer(target, lineterminator='\r\n')
    for row in sheet:
        if any(cell.data_type == 'f' for cell in row):
            raise GuardError('CSV export of formula cells is unsupported; caches are not trusted')
        writer.writerow([_csv_value(cell.value) for cell in row])
    workbook.close()
    return target.getvalue()


def make_plan(editor, request):
    if isinstance(request, (ConstrainedDocx, StructuredDocx)):
        from .professional_edits import make_plan as make_structured_plan
        return make_structured_plan(editor, request)
    source = editor.space.file(request.path)
    expected = '.pptx' if isinstance(request, (PptxReplace, PptxSlide)) else '.xlsx' if isinstance(request, ExcelRecalculate) else '.docx'
    if isinstance(request, DocumentExport):
        expected = '.docx' if request.format == 'pdf' else '.xlsx'
    if source.suffix.lower() != expected:
        raise GuardError(f'{expected} required')
    parts = _parts(source)
    from .documents import inspect_file
    inspection = inspect_file(source)
    if inspection.get('protected') or inspection.get('macros') or inspection.get('external_links'):
        raise GuardError('protected or active documents are read-only')
    before = {'source_sha256': digest(source)}
    after = request.model_dump(exclude={'path', 'kind'})
    if isinstance(request, (WordNumberedList, WordLocalImage, WordToc)):
        root = _word(parts)
        if isinstance(request, WordNumberedList):
            if 'word/numbering.xml' not in parts:
                raise GuardError('existing numbering part required')
            relations = ET.fromstring(parts['word/_rels/document.xml.rels'])
            if not any(node.get('Type', '').endswith('/numbering') and node.get('Target') in {'numbering.xml', '/word/numbering.xml'} for node in relations):
                raise GuardError('existing numbering relationship required')
            if any(not item or len(item) > 1000 or any(ord(c) < 32 and c not in '\t\n\r' for c in item) for item in request.items):
                raise GuardError('1-20 nonempty list items of at most 1000 characters required')
        if isinstance(request, WordLocalImage):
            after['image_sha256'] = digest(_image(editor.space, source, request))
        if isinstance(request, WordToc):
            fields = [n for n in root.iter(W+'fldSimple')]
            if len(fields) > 1 or any((n.get(W+'instr') or '').strip() != TOC for n in fields) or list(root.iter(W+'instrText')):
                raise GuardError('only zero or one simple standard TOC field supported')
            after['notice'] = 'TOC field rebuilt; page numbers and field cache not refreshed or verified'
    elif isinstance(request, (PptxReplace, PptxSlide)):
        slides = _slides(source, parts)
        if request.slide_index >= len(slides):
            raise GuardError('slide index out of bounds')
        if isinstance(request, PptxReplace):
            slide = ET.fromstring(parts[slides[request.slide_index]])
            matches = [shape for shape in slide.iter(P+'sp') if any(n.get('id') == str(request.shape_id) for n in shape.iter(P+'cNvPr'))]
            if len(matches) != 1:
                raise GuardError('unique simple text shape required')
            texts = list(matches[0].iter(A+'t'))
            if len(texts) != 1 or texts[0].text != request.old or request.old == request.new:
                raise GuardError('single run exact original text required')
        elif request.action == 'delete' and len(slides) <= 1:
            raise GuardError('cannot delete final slide')
        elif request.action == 'add_blank' and len(slides) >= 100:
            raise GuardError('slide limit reached')
        before['slide_count'] = len(slides)
    elif isinstance(request, ExcelRecalculate):
        after['verified_expected_values'] = formula_results(source)
    elif request.format == 'csv':
        after['csv_sha256'] = __import__('hashlib').sha256(csv_text(source, request.sheet).encode('utf-8-sig')).hexdigest()
    elif request.sheet is not None:
        raise GuardError('sheet only applies to CSV export')
    plan = Plan(uuid.uuid4().hex, request, before['source_sha256'], request.kind,
                json.dumps(before, ensure_ascii=False), json.dumps(after, ensure_ascii=False),
                ('independent confirmation required',), True)
    editor.plans[plan.id] = plan
    return plan


def render_gate(root, source, output, request=None, approved_anchors=None):
    from .fidelity import compare_documents
    try:
        if isinstance(request, RangeDocx):
            from .range_alignment import anchors_for
            if approved_anchors is not None:
                if type(approved_anchors) is not list or not 1 <= len(approved_anchors) <= 12:
                    raise GuardError('professional_capacity')
                anchors = approved_anchors
            elif request.kind == 'block_docx':
                from .block_edits import render_anchors
                anchors = render_anchors(_parts(source, range_edit=True), request)
            else:
                anchors = anchors_for(_parts(source, range_edit=True), request)
            result = compare_documents(root, source, output, static_docx=True, range_anchors=anchors)
        else:
            result = compare_documents(root, source, output)
    except Exception:
        error = GuardError('render validation failed; output publication refused')
        error.render_diff = {'render': {'status': 'error', 'equivalent': None, 'code': 'render_validation_failed'}}
        raise error from None
    result['policy'] = 'Render differences are observations; target and non-target structure checks authorize edits. Rendering errors refuse publication.'
    valid = result.get('inputs_unchanged') is True and result.get('render', {}).get('status') in {'equivalent', 'different'}
    render = result.get('render', {})
    if valid and isinstance(request, PptxReplace):
        valid = render.get('page_count_equal') is True and all(
            page.get('page') == request.slide_index+1 or
            all(page.get(key) is True for key in ('text_equal', 'size_equal', 'pixels_equal'))
            for page in render.get('pages', []))
        result['expected_render_scope'] = 'Only the selected slide may differ; page count unchanged'
    if valid and isinstance(request, PptxSlide):
        expected_delta = 1 if request.action == 'add_blank' else -1
        valid = render.get('output_pages', 0) - render.get('source_pages', 0) == expected_delta
        result['expected_render_scope'] = 'Slide count changes by exactly the planned delta; package projections protect other slides'
    if valid and isinstance(request, ExcelRecalculate):
        valid = render.get('equivalent') is True
        result['expected_render_scope'] = 'Recalculated numeric caches must not change LibreOffice rendering'
    if not valid:
        error = GuardError('render validation failed; output publication refused')
        error.render_diff = result
        raise error
    return result


def _word_write(parts, request, source, editor):
    if isinstance(request, WordNumberedList):
        root = ET.fromstring(parts['word/document.xml'])
        body = root.find(W+'body')
        numbering = ET.fromstring(parts.get('word/numbering.xml', b'<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'))
        if 'word/numbering.xml' not in parts:
            raise GuardError('existing numbering part required')
        abstract_id = max([int(n.get(W+'abstractNumId')) for n in numbering.findall(W+'abstractNum')] + [-1])+1
        num_id = max([int(n.get(W+'numId')) for n in numbering.findall(W+'num')] + [0])+1
        abstract = ET.SubElement(numbering, W+'abstractNum', {W+'abstractNumId': str(abstract_id)})
        level = ET.SubElement(abstract, W+'lvl', {W+'ilvl': '0'})
        for tag, value in [('start', '1'), ('numFmt', 'decimal'), ('lvlText', '%1.'), ('lvlJc', 'left')]:
            ET.SubElement(level, W+tag, {W+'val': value})
        num = ET.SubElement(numbering, W+'num', {W+'numId': str(num_id)})
        ET.SubElement(num, W+'abstractNumId', {W+'val': str(abstract_id)})
        for item in request.items:
            paragraph = ET.Element(W+'p')
            props = ET.SubElement(paragraph, W+'pPr')
            num_props = ET.SubElement(props, W+'numPr')
            ET.SubElement(num_props, W+'ilvl', {W+'val': '0'})
            ET.SubElement(num_props, W+'numId', {W+'val': str(num_id)})
            run = ET.SubElement(paragraph, W+'r')
            ET.SubElement(run, W+'t', {'{http://www.w3.org/XML/1998/namespace}space': 'preserve'}).text = item
            body.insert(len(body)-int(body[-1].tag == W+'sectPr'), paragraph)
        parts['word/document.xml'], parts['word/numbering.xml'] = _xml(root), _xml(numbering)
    elif isinstance(request, WordToc):
        root = ET.fromstring(parts['word/document.xml'])
        body = root.find(W+'body')
        existing = list(root.iter(W+'fldSimple'))
        if existing:
            field = existing[0]
            for child in list(field):
                field.remove(child)
        else:
            paragraph = ET.Element(W+'p')
            field = ET.SubElement(paragraph, W+'fldSimple')
            body.insert(0, paragraph)
        field.set(W+'instr', TOC)
        field.set(W+'dirty', 'true')
        run = ET.SubElement(field, W+'r')
        ET.SubElement(run, W+'t').text = '目录待更新：请在 Word 中更新域'
        parts['word/document.xml'] = _xml(root)
    else:
        document = Document(source)
        original = editor.space.file(request.path)
        image = _image(editor.space, original, request)
        document.add_picture(str(image), width=Inches(request.width_inches))
        stream = BytesIO()
        document.save(stream)
        with zipfile.ZipFile(stream) as archive:
            allowed = {'word/document.xml', 'word/_rels/document.xml.rels', '[Content_Types].xml'}
            for name in archive.namelist():
                if name in allowed or name.startswith('word/media/') and name not in parts:
                    parts[name] = archive.read(name)


def _pptx_write(parts, request, source):
    slides = _slides(source, parts)
    name = slides[request.slide_index]
    if isinstance(request, PptxReplace):
        root = ET.fromstring(parts[name])
        shape = next(shape for shape in root.iter(P+'sp') if any(n.get('id') == str(request.shape_id) for n in shape.iter(P+'cNvPr')))
        next(shape.iter(A+'t')).text = request.new
        parts[name] = _xml(root)
        return
    presentation = ET.fromstring(parts['ppt/presentation.xml'])
    ids = presentation.find(P+'sldIdLst')
    relations = ET.fromstring(parts['ppt/_rels/presentation.xml.rels'])
    types = ET.fromstring(parts['[Content_Types].xml'])
    rel_name = name.rsplit('/', 1)[0]+'/_rels/'+name.rsplit('/', 1)[1]+'.rels'
    if request.action == 'delete':
        slide_id = list(ids)[request.slide_index]
        relation = next(node for node in relations if node.get('Id') == slide_id.get(R+'id'))
        ids.remove(slide_id)
        relations.remove(relation)
        parts.pop(name)
        parts.pop(rel_name, None)
        for node in list(types):
            if node.get('PartName') == '/'+name:
                types.remove(node)
    else:
        number = max([int(m[1]) for part in parts if (m := re.fullmatch(r'ppt/slides/slide([0-9]+)\.xml', part))])+1
        new_name = f'ppt/slides/slide{number}.xml'
        rid = 'rId'+str(max([int(n.get('Id')[3:]) for n in relations if re.fullmatch('rId[0-9]+', n.get('Id', ''))]+[0])+1)
        node = ET.Element(P+'sldId', {'id': str(max(int(n.get('id')) for n in ids)+1), R+'id': rid})
        ids.insert(request.slide_index+1, node)
        ET.SubElement(relations, REL+'Relationship', {'Id': rid, 'Type': R[1:-1]+'/slide', 'Target': f'slides/slide{number}.xml'})
        ET.SubElement(types, CT+'Override', {'PartName': '/'+new_name, 'ContentType': 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'})
        root = ET.fromstring(parts[name])
        tree = root.find('./'+P+'cSld/'+P+'spTree')
        for child in list(tree):
            if child.tag not in {P+'nvGrpSpPr', P+'grpSpPr'}:
                tree.remove(child)
        parts[new_name] = _xml(root)
        if rel_name in parts:
            parts[f'ppt/slides/_rels/slide{number}.xml.rels'] = parts[rel_name]
    parts['ppt/presentation.xml'] = _xml(presentation)
    parts['ppt/_rels/presentation.xml.rels'] = _xml(relations)
    parts['[Content_Types].xml'] = _xml(types)


def _native(source, mode, extension, runs):
    from .fidelity import _temporary_root, _profile, _run, runtime_path, bounded_bytes, package_projection
    with tempfile.TemporaryDirectory(prefix='oaw-', dir=_temporary_root()) as raw:
        root = Path(raw)
        copy = root / ('source'+source.suffix.lower())
        shutil.copy2(source, copy)
        destination = root/'out'
        destination.mkdir()
        _profile(destination/'profile')
        _run((sys.executable, str(Path(__file__).with_name('fidelity_worker.py')), mode, str(copy), str(destination), str(runtime_path())), 30000, runs)
        result = destination/('source'+extension)
        if extension == '.xlsx':
            package_projection(result, for_render=True)
        return bounded_bytes(result)


def _recalculate(parts, source, expected, runs):
    converted = _native(source, '--recalculate', '.xlsx', runs)
    with zipfile.ZipFile(BytesIO(converted)) as archive:
        values_book = load_workbook(BytesIO(converted), data_only=True)
        formula_book = load_workbook(source, data_only=False)
        workbook_xml = ET.fromstring(parts['xl/workbook.xml'])
        relations = ET.fromstring(parts['xl/_rels/workbook.xml.rels'])
        import posixpath
        sheet_parts = {}
        for node in workbook_xml.find(S+'sheets'):
            target = next(n.get('Target') for n in relations if n.get('Id') == node.get(R+'id'))
            sheet_parts[node.get('name')] = posixpath.normpath(posixpath.join('xl', target)) if not target.startswith('/') else target[1:]
        for index, sheet in enumerate(formula_book, 1):
            if sheet.title not in expected:
                continue
            for address, expected_value in expected[sheet.title].items():
                actual = values_book[sheet.title][address].value
                if type(actual) not in {int, float} or not math.isclose(actual, expected_value, rel_tol=1e-10, abs_tol=1e-10):
                    raise GuardError('recalculated value differs from independent expected result')
            name = sheet_parts[sheet.title]
            original = ET.fromstring(parts[name])
            for cell in original.iter(S+'c'):
                if cell.get('r') not in expected[sheet.title]:
                    continue
                if cell.find(S+'f') is None or cell.find(S+'f').attrib or cell.get('t') not in {None, 'n'}:
                    raise GuardError('only ordinary numeric formula cells supported')
                for value in list(cell.findall(S+'v')):
                    cell.remove(value)
                ET.SubElement(cell, S+'v').text = str(values_book[sheet.title][cell.get('r')].value)
            parts[name] = _xml(original)
        values_book.close()
        formula_book.close()


def execute(editor, plan, confirmed):
    if confirmed is not True:
        raise GuardError('explicit human confirmation required')
    editor.plans.pop(plan.id)
    source = editor.space.file(plan.request.path)
    if digest(source) != plan.source_sha256:
        raise GuardError('source changed since preview')
    refreshed = make_plan(editor, plan.request)
    editor.plans.pop(refreshed.id)
    if (refreshed.before, refreshed.after) != (plan.before, plan.after):
        raise GuardError('inputs changed since preview')
    request = plan.request
    extension = '.'+request.format if isinstance(request, DocumentExport) else source.suffix
    output = editor.space.file(str(source.with_name(f'{source.stem}.agent-{plan.id[:8]}{extension}')), existing=False)
    editor.audit.mkdir(exist_ok=True)
    from .writes import _audit_path
    _audit_path(editor.audit, directory=True)
    job = editor.audit/plan.id
    job.mkdir()
    backup, working = job/('backup'+source.suffix), job/('working'+extension)
    log = {**plan.public(), 'status': 'failed', 'backup': str(backup), 'output': str(output), 'processes': []}
    published = False
    try:
        shutil.copy2(source, backup)
        if digest(backup) != plan.source_sha256:
            raise GuardError('backup mismatch')
        parts = _parts(backup, range_edit=isinstance(request, RangeDocx))
        before_parts = dict(parts)
        expected = json.loads(plan.after)
        if isinstance(request, DocumentExport):
            if request.format == 'csv':
                payload = csv_text(backup, request.sheet).encode('utf-8-sig')
                working.write_bytes(payload)
                if __import__('hashlib').sha256(working.read_bytes()).hexdigest() != expected['csv_sha256']:
                    raise GuardError('CSV export verification failed')
                with working.open(encoding='utf-8-sig', newline='') as stream:
                    rows = list(csv.reader(stream))
                log['export_verification'] = {'rows': len(rows), 'sheet': request.sheet, 'scope': 'literal values only; formatting intentionally omitted'}
            else:
                working.write_bytes(_native(backup, '--convert', '.pdf', log['processes']))
                from .fidelity import _run, _temporary_root, bounded_bytes
                with tempfile.TemporaryDirectory(prefix='oap-', dir=_temporary_root()) as raw:
                    result = Path(raw)/'pdf.json'
                    _run((sys.executable, str(Path(__file__).with_name('fidelity_worker.py')), str(working), str(working), str(result)), 30000, log['processes'])
                    log['export_verification'] = json.loads(bounded_bytes(result))
                    if log['export_verification'].get('equivalent') is not True:
                        raise GuardError('PDF reopening failed')
        else:
            if isinstance(request, (ConstrainedDocx, StructuredDocx)):
                from .professional_edits import transform
                updated, _, _ = transform(parts, request)
                parts.clear()
                parts.update(updated)
            elif isinstance(request, (WordNumberedList, WordLocalImage, WordToc)):
                _word_write(parts, request, backup, editor)
            elif isinstance(request, (PptxReplace, PptxSlide)):
                _pptx_write(parts, request, backup)
            else:
                _recalculate(parts, backup, expected['verified_expected_values'], log['processes'])
                log['calculation'] = {'status': 'independently_verified', 'values': expected['verified_expected_values'], 'engine': 'LibreOffice numeric cache; only verified v nodes transplanted'}
            verify_change(before_parts, parts, request)
            if isinstance(request, WordLocalImage):
                image = _image(editor.space, source, request)
                if digest(image) != expected['image_sha256']:
                    raise GuardError('image changed during execution')
            _zip(parts, working)
            reopened = _parts(working, range_edit=isinstance(request, RangeDocx))
            if reopened != parts:
                raise GuardError('reopened package differs from staged parts')
            if source.suffix.lower() == '.docx':
                reopened_document = Document(working)
                if isinstance(request, WordLocalImage):
                    from hashlib import sha256
                    shape = reopened_document.inline_shapes[-1]
                    rid = shape._inline.graphic.graphicData.pic.blipFill.blip.embed
                    if sha256(reopened_document.part.related_parts[rid].blob).hexdigest() != expected['image_sha256']:
                        raise GuardError('embedded image differs from confirmed image')
                    if shape.width != Inches(request.width_inches):
                        raise GuardError('embedded image width differs from preview')
            elif source.suffix.lower() == '.pptx':
                from .presentations import _load
                _load(working)
            else:
                cached = load_workbook(working, data_only=True)
                for sheet, values in expected['verified_expected_values'].items():
                    for cell, value in values.items():
                        if not math.isclose(cached[sheet][cell].value, value, rel_tol=1e-10, abs_tol=1e-10):
                            raise GuardError('published cache verification failed')
                cached.close()
            log['render_diff'] = render_gate(editor.space.root, backup, working, request)
        if editor.space.file(str(source)) != source or digest(source) != plan.source_sha256 or digest(backup) != plan.source_sha256:
            raise GuardError('source or backup changed during execution')
        os.link(working, output)
        published = True
        log.update(status='verified', output_sha256=digest(output), diff={'target': plan.target, 'before': plan.before, 'after': plan.after})
        (job/'audit.json').write_text(json.dumps(log, ensure_ascii=False, indent=2), encoding='utf-8')
        return log
    except Exception as exc:
        log['status'] = 'failed'
        if hasattr(exc, 'render_diff'):
            log['render_diff'] = exc.render_diff
        if published:
            output.unlink(missing_ok=True)
        (job/'audit.json').write_text(json.dumps(log, ensure_ascii=False, indent=2), encoding='utf-8')
        raise
    finally:
        working.unlink(missing_ok=True)

def _canonical(node):
    return ET.tostring(node, method='c14n')


def verify_change(before, after, request):
    if isinstance(request, (ConstrainedDocx, StructuredDocx)):
        from .professional_edits import verify
        return verify(before, after, request)
    changed = {name for name in before.keys() | after.keys() if before.get(name) != after.get(name)}
    if isinstance(request, (WordNumberedList, WordLocalImage, WordToc)):
        allowed = {'word/document.xml'}
        if isinstance(request, WordNumberedList):
            allowed.add('word/numbering.xml')
        if isinstance(request, WordLocalImage):
            allowed |= {'word/_rels/document.xml.rels', '[Content_Types].xml'}
            allowed |= {name for name in after if name.startswith('word/media/') and name not in before}
        if not changed <= allowed or not before.keys() <= after.keys():
            raise GuardError('non-target Word package part changed')
        original, result = (ET.fromstring(parts['word/document.xml']) for parts in (before, after))
        old_body, new_body = original.find(W+'body'), result.find(W+'body')
        if isinstance(request, WordToc):
            old_fields, new_fields = list(old_body.iter(W+'fldSimple')), list(new_body.iter(W+'fldSimple'))
            if len(new_fields) != 1 or new_fields[0].get(W+'instr') != TOC:
                raise GuardError('TOC target verification failed')
            if old_fields:
                field = old_fields[0]
                field.getparent().replace(field, deepcopy(new_fields[0]))
            else:
                added = new_body[0]
                if len(added) != 1 or added[0] is not new_fields[0]:
                    raise GuardError('unexpected TOC insertion')
                new_body.remove(added)
        else:
            added_count = len(request.items) if isinstance(request, WordNumberedList) else 1
            new_section = new_body.find(W+'sectPr')
            end = len(new_body) - int(new_section is not None)
            additions = list(new_body)[end-added_count:end]
            if len(new_body) != len(old_body)+added_count or any(node.tag != W+'p' for node in additions):
                raise GuardError('unexpected body insertion')
            if isinstance(request, WordNumberedList):
                if [''.join(node.itertext()) for node in additions] != request.items:
                    raise GuardError('numbered list text mismatch')
                old_num, new_num = (ET.fromstring(parts['word/numbering.xml']) for parts in (before, after))
                if len(new_num) != len(old_num)+2 or [_canonical(n) for n in old_num] != [_canonical(n) for n in list(new_num)[:-2]]:
                    raise GuardError('existing numbering changed')
                if [n.tag for n in list(new_num)[-2:]] != [W+'abstractNum', W+'num']:
                    raise GuardError('new numbering definition missing')
                num_id = new_num[-1].get(W+'numId')
                if any(node.find('./'+W+'pPr/'+W+'numPr/'+W+'numId').get(W+'val') != num_id for node in additions):
                    raise GuardError('numbered list reference mismatch')
            elif len(list(additions[0].iter(W+'drawing'))) != 1:
                raise GuardError('image target missing')
            for node in additions:
                new_body.remove(node)
        if _canonical(original) != _canonical(result):
            raise GuardError('non-target Word document content changed')
        if isinstance(request, WordLocalImage):
            for name in ('word/_rels/document.xml.rels', '[Content_Types].xml'):
                old, new = ET.fromstring(before[name]), ET.fromstring(after[name])
                old_nodes = {_canonical(node) for node in old}
                new_nodes = {_canonical(node) for node in new}
                if not old_nodes <= new_nodes:
                    raise GuardError('existing image relationships or content types changed')
    elif isinstance(request, ExcelRecalculate):
        allowed = {name for name in before if re.fullmatch(r'xl/worksheets/[^/]+\.xml', name)}
        if not changed <= allowed or before.keys() != after.keys():
            raise GuardError('non-target workbook part changed')
        for name in changed:
            old, new = ET.fromstring(before[name]), ET.fromstring(after[name])
            for tree in (old, new):
                for cell in tree.iter(S+'c'):
                    if cell.find(S+'f') is not None:
                        for value in list(cell.findall(S+'v')):
                            cell.remove(value)
            if _canonical(old) != _canonical(new):
                raise GuardError('recalculation changed content outside numeric caches')
    else:
        from .presentations import _load
        if isinstance(request, PptxReplace):
            slides = ET.fromstring(before['ppt/presentation.xml'])
            ids = slides.find(P+'sldIdLst')
            relations = ET.fromstring(before['ppt/_rels/presentation.xml.rels'])
            rid = list(ids)[request.slide_index].get(R+'id')
            target = next(n.get('Target') for n in relations if n.get('Id') == rid)
            import posixpath
            name = posixpath.normpath(posixpath.join('ppt', target)) if not target.startswith('/') else target[1:]
            if changed != {name}:
                raise GuardError('non-target slide package part changed')
            old, new = ET.fromstring(before[name]), ET.fromstring(after[name])
            shape = next(s for s in old.iter(P+'sp') if any(n.get('id') == str(request.shape_id) for n in s.iter(P+'cNvPr')))
            next(shape.iter(A+'t')).text = request.new
            if _canonical(old) != _canonical(new):
                raise GuardError('non-target slide content changed')
        else:
            allowed = {'ppt/presentation.xml', 'ppt/_rels/presentation.xml.rels', '[Content_Types].xml'}
            allowed |= {name for name in before.keys() ^ after.keys() if name.startswith('ppt/slides/')}
            if not changed <= allowed:
                raise GuardError('existing slide content changed')
            old = ET.fromstring(before['ppt/presentation.xml'])
            new = ET.fromstring(after['ppt/presentation.xml'])
            old_ids, new_ids = old.find(P+'sldIdLst'), new.find(P+'sldIdLst')
            delta = 1 if request.action == 'add_blank' else -1
            if len(new_ids) != len(old_ids)+delta:
                raise GuardError('slide count mismatch')
            if delta == 1:
                new_ids.remove(new_ids[request.slide_index+1])
            else:
                old_ids.remove(old_ids[request.slide_index])
            if _canonical(old) != _canonical(new):
                raise GuardError('non-target presentation metadata changed')
