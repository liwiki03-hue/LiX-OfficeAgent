from copy import deepcopy
import hashlib
import io
import json
import os
from pathlib import Path, PurePosixPath
import posixpath
import re
import shutil
import stat
from threading import Event
import time
from urllib.parse import urlsplit
import uuid
import zipfile

from docx import Document
from lxml import etree as ET

from .writes import GuardError, digest
from .edit_metrics import publication_metrics


W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
R = '{http://schemas.openxmlformats.org/package/2006/relationships}'
MAX_BYTES = 8 * 1024 * 1024
MAX_EXPANDED = 64 * 1024 * 1024
MAX_PART = 16 * 1024 * 1024
MAX_PARAGRAPHS = 2000


def sha(data):
    return hashlib.sha256(data).hexdigest()


def xml(data):
    parser = ET.XMLParser(resolve_entities=False, load_dtd=False, no_network=True, huge_tree=False)
    root = ET.fromstring(data, parser)
    if root.getroottree().docinfo.doctype or any(isinstance(n, ET._Entity) for n in root.iter()):
        raise GuardError('拒绝XML实体或文档类型声明')
    if sum(1 for _ in root.iter()) > 200000:
        raise GuardError('文档结构超过编辑预算，请拆分材料')
    return root


def parts(data):
    if not isinstance(data, bytes) or not 1 <= len(data) <= MAX_BYTES:
        raise GuardError('分页编辑首版仅接收8 MiB以内的DOCX快照')
    result, seen, total = {}, set(), 0
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            entries = archive.infolist()
            if not 1 <= len(entries) <= 2000:
                raise GuardError('文档包部件超过编辑预算')
            for entry in entries:
                name, path = entry.filename, PurePosixPath(entry.filename)
                total += entry.file_size
                if (not name or name in seen or '\\' in name or ':' in name or path.is_absolute() or '..' in path.parts or
                        stat.S_ISLNK(entry.external_attr >> 16) or entry.flag_bits & 1 or
                        entry.file_size > MAX_PART or total > MAX_EXPANDED or
                        entry.compress_type not in {zipfile.ZIP_DEFLATED, zipfile.ZIP_STORED}):
                    raise GuardError('文档包包含不安全或超限部件')
                seen.add(name)
                if entry.is_dir():
                    if entry.file_size or name.rstrip('/') != str(path):
                        raise GuardError('文档包目录标记无效')
                    continue
                if name != str(path) or entry.file_size > 1000 * max(1, entry.compress_size):
                    raise GuardError('文档包路径或压缩比异常')
                value = archive.read(entry)
                if name.lower().endswith(('.xml', '.rels')):
                    xml(value)
                if any(token in name.lower() for token in ['vbaproject', 'activex', 'embeddings/', 'externallinks/', 'connections.xml']):
                    raise GuardError('分页编辑不接收宏、嵌入对象、外链数据或活动内容')
                result[name] = value
    except (zipfile.BadZipFile, ET.XMLSyntaxError, RuntimeError, OSError) as error:
        raise GuardError('DOCX包损坏或无法解析') from error
    if not {'[Content_Types].xml', '_rels/.rels', 'word/document.xml', 'word/styles.xml'} <= result.keys():
        raise GuardError('缺少DOCX必要部件')
    types = xml(result['[Content_Types].xml'])
    expected = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'
    if not any(n.get('PartName') == '/word/document.xml' and n.get('ContentType') == expected for n in types):
        raise GuardError('主部件不是普通DOCX')
    if any('macro' in n.get('ContentType', '').lower() for n in types):
        raise GuardError('拒绝宏内容类型')
    for name, value in result.items():
        if not name.endswith('.rels'):
            continue
        identifiers = set()
        for rel in xml(value):
            identifier, target = rel.get('Id'), rel.get('Target', '')
            if rel.tag != R+'Relationship' or not identifier or identifier in identifiers:
                raise GuardError('DOCX关系标识无效')
            identifiers.add(identifier)
            if rel.get('TargetMode', 'Internal') != 'Internal':
                parsed = urlsplit(target)
                if not rel.get('Type', '').endswith('/hyperlink') or parsed.scheme not in {'http', 'https', 'mailto'} or parsed.username or parsed.password:
                    raise GuardError('文档包含不支持的外部关系')
                continue
            base = posixpath.dirname(posixpath.dirname(name)) if name != '_rels/.rels' else ''
            resolved = posixpath.normpath(posixpath.join(base, target)) if not target.startswith('/') else target[1:]
            if not target or '\\' in target or ':' in target or resolved.startswith('../') or resolved not in result:
                raise GuardError('DOCX内部关系目标缺失或越界')
    return result


def normalize(data, original=None):
    values = parts(data)
    if original is not None:
        source = parts(original)
        for name,value in source.items():
            if name.startswith(('customXml/', 'docProps/', 'word/theme/')) or name in {'_rels/.rels','word/stylesWithEffects.xml','word/fontTable.xml','word/settings.xml','word/webSettings.xml'}:
                values[name] = value
        for name in list(values):
            if name.startswith('docProps/') and name not in source:
                del values[name]
        types = xml(values['[Content_Types].xml'])
        for node in list(types):
            target = node.get('PartName','').lstrip('/')
            if target and target not in values:
                types.remove(node)
        values['[Content_Types].xml'] = ET.tostring(types,xml_declaration=True,encoding='UTF-8',standalone=True)
        _restore_unchanged_inherited_colors(values, source)
    stream = io.BytesIO()
    with zipfile.ZipFile(stream, 'w', zipfile.ZIP_DEFLATED) as archive:
        for name, value in values.items():
            archive.writestr(name, value)
    return stream.getvalue()


def signature(node):
    if node is None:
        return None
    attributes = sorted((key, value) for key, value in node.attrib.items() if not key.split('}')[-1].startswith('rsid') and key not in {
        '{http://schemas.microsoft.com/office/word/2010/wordml}paraId', '{http://schemas.microsoft.com/office/word/2010/wordml}textId'})
    text = node.text if not len(node) and node.text and (node.text.strip() or node.tag == W+'t') else None
    return [node.tag, attributes, text, [signature(child) for child in node]]


def _restore_unchanged_inherited_colors(values, source):
    if signature(xml(values['word/styles.xml'])) != signature(xml(source['word/styles.xml'])):
        return
    theme_parts = {name: value for name, value in source.items() if name.startswith('word/theme/')}
    if theme_parts != {name: value for name, value in values.items() if name.startswith('word/theme/')}:
        return
    styles = xml(source['word/styles.xml'])
    style_map = {node.get(W+'styleId'): node for node in styles.findall(W+'style')}
    default_style = next((key for key, node in style_map.items() if node.get(W+'type') == 'paragraph' and node.get(W+'default') == '1'), None)
    defaults = styles.find('./'+W+'docDefaults/'+W+'rPrDefault/'+W+'rPr/'+W+'color')
    theme_colors = {}
    if len(theme_parts) == 1:
        drawing = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
        scheme = xml(next(iter(theme_parts.values()))).find('.//'+drawing+'clrScheme')
        if scheme is not None:
            theme_colors = {ET.QName(node).localname: node[0].get('val') for node in scheme if len(node) == 1 and node[0].tag == drawing+'srgbClr'}

    def inherited_color(paragraph, run):
        color = defaults
        paragraph_style = paragraph.find('./'+W+'pPr/'+W+'pStyle')
        character_style = run.find('./'+W+'rPr/'+W+'rStyle')
        for style_id in [paragraph_style.get(W+'val') if paragraph_style is not None else default_style, character_style.get(W+'val') if character_style is not None else None]:
            chain, seen = [], set()
            while style_id:
                if style_id in seen or style_id not in style_map:
                    return None
                seen.add(style_id)
                node = style_map[style_id]
                chain.append(node)
                parent = node.find(W+'basedOn')
                style_id = parent.get(W+'val') if parent is not None else None
            for node in reversed(chain):
                override = node.find('./'+W+'rPr/'+W+'color')
                if override is not None:
                    color = override
        if color is None or set(color.attrib) - {W+'val', W+'themeColor'}:
            return None
        value, theme = color.get(W+'val', ''), color.get(W+'themeColor')
        if not re.fullmatch('[0-9A-Fa-f]{6}', value) or theme and (theme_colors.get(theme) or '').upper() != value.upper():
            return None
        return value.upper()

    before, after = xml(source['word/document.xml']), xml(values['word/document.xml'])
    old_paragraphs, new_paragraphs = list(before.iter(W+'p')), list(after.iter(W+'p'))
    if len(old_paragraphs) != len(new_paragraphs):
        return
    changed = False
    for old, new in zip(old_paragraphs, new_paragraphs):
        old_runs, new_runs = old.findall(W+'r'), new.findall(W+'r')
        if len(old_runs) != len(new_runs):
            continue
        removed = []
        for old_run, new_run in zip(old_runs, new_runs):
            properties = new_run.find(W+'rPr')
            color = properties.find(W+'color') if properties is not None else None
            if color is not None and old_run.find('./'+W+'rPr/'+W+'color') is None and set(color.attrib) == {W+'val'} and not len(color) and color.get(W+'val', '').upper() == inherited_color(old, old_run):
                removed.append((properties, properties.index(color), color))
                properties.remove(color)
        if removed and signature(old) == signature(new):
            changed = True
        else:
            for properties, index, color in removed:
                properties.insert(index, color)
    if changed:
        values['word/document.xml'] = ET.tostring(after, xml_declaration=True, encoding='UTF-8', standalone=True)


def projection(data):
    book = Document(io.BytesIO(data))
    if len(book.paragraphs) > MAX_PARAGRAPHS:
        raise GuardError('正文超过2000段编辑预算，请拆分材料')
    style_map = {s.style_id:s.element for s in book.styles}
    defaults = book.styles.element.find(W+'docDefaults')
    language = defaults.find('.//'+W+'lang') if defaults is not None else None
    default_language = dict(language.attrib) if language is not None else {}
    def inherited_language(style_id, initial):
        chain, seen = [],set()
        while style_id in style_map and style_id not in seen:
            seen.add(style_id);node=style_map[style_id];chain.append(node)
            parent=node.find(W+'basedOn');style_id=parent.get(W+'val') if parent is not None else None
        result=dict(initial)
        for node in reversed(chain):
            item=node.find('./'+W+'rPr/'+W+'lang')
            if item is not None:result.update(item.attrib)
        return result
    def paragraph(p):
        from docx.text.paragraph import Paragraph
        node=deepcopy(p._p)
        inherited=inherited_language(p.style.style_id,default_language)
        previous=None
        for child in list(node):
            if child.tag!=W+'r':previous=None;continue
            properties=child.find(W+'rPr')
            if properties is not None:
                char_style=properties.find(W+'rStyle')
                effective=inherited_language(char_style.get(W+'val') if char_style is not None else None,inherited)
                lang=properties.find(W+'lang')
                if lang is not None and all(effective.get(k)==v for k,v in lang.attrib.items()):properties.remove(lang)
                if not len(properties) and not properties.attrib:child.remove(properties)
            if previous is not None and signature(previous.find(W+'rPr'))==signature(child.find(W+'rPr')) and signature(previous)[1]==signature(child)[1]:
                for item in list(child):
                    if item.tag!=W+'rPr':previous.append(item)
                node.remove(child)
            else:previous=child
        for run in node.findall(W+'r'):
            previous=None
            for child in list(run):
                if child.tag!=W+'t':previous=None;continue
                child.attrib.pop('{http://www.w3.org/XML/1998/namespace}space',None)
                if previous is not None and previous.attrib==child.attrib:
                    previous.text=(previous.text or '')+(child.text or '');run.remove(child)
                else:previous=child
        normalized=Paragraph(node,p._parent)
        return {'text': p.text, 'style': p.style.style_id, 'properties': signature(p._p.pPr),
                'inline':signature(node),
                'runs': [{'text': run.text, 'properties': signature(run._r.rPr)} for run in normalized.runs]}
    values = parts(data)
    used = {s.style_id for s in book.styles if s.element.get(W+'default') == '1'}
    for name,value in values.items():
        if name.startswith('word/') and name.endswith('.xml') and name != 'word/styles.xml':
            node = xml(value)
            used.update(n.get(W+'val') for tag in ['pStyle','rStyle','tblStyle'] for n in node.iter(W+tag) if n.get(W+'val'))
    while True:
        related = {n.get(W+'val') for key in used if key in style_map for tag in ['basedOn','link','next'] for n in style_map[key].findall(W+tag) if n.get(W+'val')}
        if related <= used:
            break
        used.update(related)
    return {'paragraphs': [paragraph(p) for p in book.paragraphs],
            'tables': [{'grid': signature(t._tbl.tblGrid), 'properties': signature(t._tbl.tblPr),
                        'rows': [[{'properties':signature(c._tc.tcPr), 'paragraphs':[paragraph(p) for p in c.paragraphs]} for c in row.cells] for row in t.rows]} for t in book.tables],
            'sections': [signature(s._sectPr) for s in book.sections],
            'headers': [[paragraph(p) for p in s.header.paragraphs] for s in book.sections],
            'footers': [[paragraph(p) for p in s.footer.paragraphs] for s in book.sections],
            'styles': {key:signature(style_map[key]) for key in used if key in style_map},
            'defaults':signature(xml(values['word/styles.xml']).find(W+'docDefaults')),
            'opaque':{name:sha(value) for name,value in values.items() if name.startswith('customXml/')},
            'media': {name:sha(value) for name,value in values.items() if name.startswith('word/media/')}}


def eligibility(data):
    values = parts(data)
    document = xml(values['word/document.xml'])
    unsupported = {'altChunk', 'object', 'sdt', 'ins', 'del', 'moveFrom', 'moveTo', 'footnoteReference', 'endnoteReference', 'commentRangeStart', 'commentRangeEnd', 'pict', 'customXml', 'framePr', 'hyperlink', 'bookmarkStart', 'bookmarkEnd'}
    features = sorted({ET.QName(n).localname for n in document.iter() if isinstance(n.tag, str) and ET.QName(n).localname in unsupported})
    if len(document.findall('.//'+W+'sectPr')) > 1:
        features.append('多节页面')
    if document.find('.//'+W+'instrText') is not None or document.find('.//'+W+'fldSimple') is not None:
        features.append('正文域（本轮只读保留）')
    if any(node.find('.//'+W+'tbl') is not None for node in document.findall('.//'+W+'tc')):
        features.append('嵌套表格（本轮先只读保留）')
    if document.find('.//'+W+'drawing') is not None:
        features.append('图片或绘图（本轮先只读保留）')
    if document.find('.//'+W+'vMerge') is not None or document.find('.//'+W+'gridSpan') is not None:
        features.append('合并表格（本轮先只读保留）')
    book = Document(io.BytesIO(data))
    if len(book.paragraphs) > MAX_PARAGRAPHS:
        features.append('超过2000段编辑预算')
    if len(document.findall('.//'+W+'p'))>3000:
        features.append('正文与表格段落总量超过3000段编辑预算')
    known = {'[Content_Types].xml','_rels/.rels','word/document.xml','word/_rels/document.xml.rels','word/styles.xml','word/stylesWithEffects.xml','word/numbering.xml','word/settings.xml','word/fontTable.xml','word/webSettings.xml'}
    unknown = [name for name in values if name not in known and not name.startswith(('docProps/','customXml/','word/theme/','word/media/')) and not re.fullmatch(r'word/(?:header|footer)\d+\.xml',name)]
    if unknown:
        features.append('未验证部件：'+ '、'.join(unknown[:6]))
    return {'editable':not features, 'reasons':features, 'styles':[{'id':s.style_id,'name':s.name} for s in book.styles if s.type == 1],
            'font_families':sorted({r.font.name for p in book.paragraphs for r in p.runs if r.font.name} | {s.font.name for s in book.styles if hasattr(s, 'font') and s.font.name}),
            'paragraphs':len(book.paragraphs), 'tables':len(book.tables)}


def changes(before, after):
    a, b = projection(before), projection(after)
    left_parts,right_parts=parts(before),parts(after)
    left_styles={n.get(W+'styleId'):signature(n) for n in xml(left_parts['word/styles.xml']).findall(W+'style')}
    right_styles={n.get(W+'styleId'):signature(n) for n in xml(right_parts['word/styles.xml']).findall(W+'style')}
    if any(right_styles.get(key)!=value for key,value in left_styles.items()) or any(key not in left_styles for key in b['styles']):
        raise GuardError('已有样式定义被改写或引入未批准样式；本轮只允许应用文档已有样式')
    output = []
    def formatting(row):
        return None if row is None else [row['style'],row['properties'],[r['properties'] for r in row['runs']]]
    def description(row):
        if row is None:return '无此段落'
        values=['样式：'+row['style']]
        for properties in [row['properties']]+[r['properties'] for r in row['runs']]:
            if not properties:continue
            for node in properties[3]:
                name=node[0].split('}')[-1];attrs={k.split('}')[-1]:v for k,v in node[1]}
                if name=='sz':values.append('直接字号：'+str(int(attrs.get('val','0'))/2)+' 磅')
                elif name=='rFonts':values.append('直接字体：'+' / '.join(sorted(set(attrs.values()))))
                elif name in {'b','i','u','color','highlight','jc'}:values.append({'b':'粗体','i':'斜体','u':'下划线','color':'文字色','highlight':'突出显示','jc':'对齐'}[name]+'：'+attrs.get('val','开启'))
                elif name=='spacing':
                    for key,label in [('before','段前'),('after','段后')]:
                        if attrs.get(key,'').lstrip('-').isdigit():values.append(label+'：'+str(int(attrs[key])/20)+' 磅')
        return '；'.join(dict.fromkeys(values))+'（未列出的字体与格式仍按文档样式继承）'
    for index in range(max(len(a['paragraphs']), len(b['paragraphs']))):
        left = a['paragraphs'][index] if index < len(a['paragraphs']) else None
        right = b['paragraphs'][index] if index < len(b['paragraphs']) else None
        if left != right:
            operation = 'insert' if left is None else 'delete' if right is None else 'edit'
            output.append({'kind':'paragraph', 'position':index+1, 'before':left['text'] if left else '', 'after':right['text'] if right else '',
                           'operation':operation, 'format_changed':left is not None and right is not None and formatting(left)!=formatting(right)})
            if output[-1]['format_changed']:output[-1].update(format_before=description(left),format_after=description(right))
    if len(a['tables'])!=len(b['tables']):
        raise GuardError('首版只编辑已有普通表格内容；表格增删未开放，请撤销这项结构操作')
    for index,(left,right) in enumerate(zip(a['tables'],b['tables'])):
        if left['grid']!=right['grid'] or left['properties']!=right['properties'] or len(left['rows'])!=len(right['rows']):
            raise GuardError('表格几何或行数发生变化，首版未开放；草稿保留，请撤销结构操作')
        for row_index,(lrow,rrow) in enumerate(zip(left['rows'],right['rows'])):
            if len(lrow)!=len(rrow):raise GuardError('表格列数改变，首版未开放')
            for column,(lcell,rcell) in enumerate(zip(lrow,rrow)):
                if lcell['properties']!=rcell['properties']:raise GuardError('单元格几何、合并或边界改变，首版未开放')
                if lcell!=rcell:
                    output.append({'kind':'table_cell','position':f'表格{index+1} · 行{row_index+1}列{column+1}',
                                   'before':'\n'.join(p['text'] for p in lcell['paragraphs']),'after':'\n'.join(p['text'] for p in rcell['paragraphs']),
                                   'format_changed':[formatting(p) for p in lcell['paragraphs']]!=[formatting(p) for p in rcell['paragraphs']]})
    for key,label in [('headers','页眉'),('footers','页脚')]:
        if a[key]!=b[key]:
            output.append({'kind':key,'position':label,'before':'\n'.join(p['text'] for group in a[key] for p in group),'after':'\n'.join(p['text'] for group in b[key] for p in group),'format_changed':True})
    if a['sections']!=b['sections']:
        raise GuardError('页面或节设置发生变化，本轮未开放；草稿保留，不能未经说明发布')
    for key, label in [('defaults','默认格式'),('opaque','文献与自定义元数据'),('media','媒体')]:
        if a[key] != b[key]:
            output.append({'kind':key, 'position':label, 'before':'原稿'+label, 'after':'草稿'+label, 'format_changed':True})
    if len(output) > 250:
        raise GuardError('本次影响超过250项，请分成多个明确副本任务')
    return output


class PublicationCancelled(GuardError):
    pass


class PublicationRejected(GuardError):
    def __init__(self, reason, text):
        super().__init__(text)
        self.reason = reason


class PairPublication:
    def __init__(self, lock, validate):
        self.lock = lock
        self.validate = validate
        self.stopped = Event()
        self.committing = False

    def cancel(self):
        with self.lock:
            if self.committing:
                return False
            self.stopped.set()
            return True

    def check(self):
        if self.stopped.is_set():
            raise PublicationCancelled('已停止校验，副本未发布。')

    def begin_commit(self):
        with self.lock:
            self.check()
            self.validate()
            self.committing = True


class CandidatePublisher:
    def __init__(self, editor, source, source_sha256):
        self.editor, self.source, self.source_sha256 = editor, source, source_sha256
        info = source.lstat()
        self.source_identity = (info.st_dev, info.st_ino)
        self.plans = {}

    def phase(self, name, text):
        metrics = getattr(self, 'metrics_trace', None)
        if metrics is not None:
            metrics.phase(name)
        callback = getattr(self, 'progress', None)
        if callback:
            callback(name, text)

    def check_source(self):
        info = self.source.lstat()
        if self.editor.space.file(str(self.source)) != self.source or digest(self.source) != self.source_sha256 or (info.st_dev, info.st_ino) != self.source_identity:
            raise GuardError('原文件已在外部改变，旧草稿发布计划失效；请重新准入并比较')

    def plan(self, candidate, *, document_id, draft_revision):
        self.check_source()
        original = self.source.read_bytes()
        raw_parts = parts(candidate)
        normalized = normalize(candidate, original)
        policy = eligibility(original)
        if not policy['editable']:
            raise GuardError('此文档本轮仅支持只读：'+ '、'.join(policy['reasons']))
        if not eligibility(normalized)['editable']:
            raise GuardError('候选产生了未支持结构，已拒绝发布')
        diff = changes(original, normalized)
        if not diff:
            raise GuardError('没有实际内容或格式变化，无需生成副本')
        identifier = uuid.uuid4().hex
        self.plans.clear()
        self.plans[identifier] = {'id':identifier, 'source':str(self.source), 'source_sha256':self.source_sha256,
                                 'document_id':document_id, 'draft_revision':draft_revision, 'candidate_sha256':sha(normalized),
                                 'normalization_parts':[name for name,value in parts(normalized).items() if raw_parts.get(name)!=value],
                                 'changes':diff, 'deadline':time.monotonic()+300, 'candidate':normalized}
        return self.public(identifier)

    def public(self, identifier):
        return {key:deepcopy(value) for key,value in self.plans[identifier].items() if key not in {'candidate','deadline'}}

    def claim(self, identifier, *, document_id, draft_revision, confirmed=False):
        if confirmed is not True:
            raise GuardError('生成副本需要独立确认')
        plan = self.plans.pop(identifier, None)
        if not plan or plan['document_id'] != document_id or plan['draft_revision'] != draft_revision or time.monotonic()>plan['deadline']:
            raise GuardError('发布计划已过期或草稿改变，请重新预览确认')
        self.check_source()
        return plan

    @publication_metrics
    def execute(self, identifier, *, document_id, draft_revision, confirmed=False, _claimed=None, _pair=None):
        plan = _claimed if _claimed is not None else self.claim(identifier, document_id=document_id, draft_revision=draft_revision, confirmed=confirmed)
        if confirmed is not True or plan['id'] != identifier or plan['document_id'] != document_id or plan['draft_revision'] != draft_revision:
            raise GuardError('冻结的发布归属不符')
        if _pair is not None and plan.get('kind', 'paged_candidate') != 'paged_candidate':
            raise GuardError('双输入校验仅适用于 Word 草稿副本')
        self.check_source()
        audit = self.editor.audit
        audit.mkdir(exist_ok=True)
        if audit.is_symlink() or not audit.resolve().is_relative_to(self.editor.space.root):
            raise GuardError('审计目录已被重定向')
        job = audit / identifier
        job.mkdir()
        backup, candidate = job/'backup.docx', job/'working.docx'
        output = self.editor.space.file(str(self.source.with_name(f'{self.source.stem}.agent-{identifier[:8]}.docx')), existing=False)
        log = {key:value for key,value in plan.items() if key not in {'candidate','deadline','draft_preview_base'}}
        log.update(kind=plan.get('kind','paged_candidate'), status='failed', backup=str(backup), output=str(output), normalization=plan.get('normalization','empty ZIP directory markers omitted; original bibliography/custom XML, document properties, themes, font table, settings and stylesWithEffects preserved because these are not editable in this workflow; unrequested generated document properties removed'))
        published = False
        try:
            if _pair is not None:
                log['render_mode'] = 'docx_pair_opt_in'
                _pair.check()
            self.phase('backup', '正在备份原件，原文件保持不变。')
            shutil.copy2(self.source, backup)
            if digest(backup) != self.source_sha256:
                raise GuardError('备份校验失败')
            self.phase('working_file', '正在写入待校验的工作文件，尚未发布副本。')
            with candidate.open('xb') as stream:
                stream.write(plan['candidate'])
                stream.flush()
                os.fsync(stream.fileno())
            if digest(candidate) != plan['candidate_sha256']:
                raise GuardError('候选快照身份不符')
            self.phase('reopening', '工作文件已生成，正在重新打开并核对内容。')
            reopened = projection(candidate.read_bytes())
            if reopened != projection(plan['candidate']):
                raise GuardError('重开结果与已批准草稿不一致')
            from .fidelity import DOCX_PAIR_VARIANT, compare_documents
            self.phase('rendering', '正在校验工作文件的内容和版式，完成后才会发布副本。')
            options = {'batch_conversion': True, 'cancelled': _pair.stopped.is_set} if _pair is not None else {}
            render = compare_documents(self.editor.space.root, backup, candidate, static_docx=True, reuse_verified=True, **options)
            log['render_diff'] = render
            rendering = render.get('render', {})
            if _pair is not None and rendering.get('code') == 'worker_cancelled' and _pair.stopped.is_set():
                _pair.check()
            if render.get('inputs_unchanged') is not True or rendering.get('status') not in {'equivalent','different'}:
                raise PublicationRejected('render_failed', '渲染闸门未通过，副本未发布')
            if _pair is not None and (render.get('batch_conversion', {}).get('used') is not True
                    or render.get('conversion_variant') != DOCX_PAIR_VARIANT
                    or render.get('source_sha256') != self.source_sha256
                    or render.get('output_sha256') != plan['candidate_sha256']
                    or render.get('render_environment', {}).get('complete') is not True):
                raise PublicationRejected('render_receipt_mismatch', '双输入校验归属不符，副本未发布')
            metrics = getattr(self, 'metrics_trace', None)
            if metrics is not None:
                metrics.phase('validation')
            self.check_source()
            if digest(candidate) != plan['candidate_sha256'] or digest(backup) != self.source_sha256:
                raise GuardError('验证期间快照或备份发生变化')
            self.phase('publishing', '内容与版式校验已完成，正在发布已验证副本。')
            if _pair is not None:
                _pair.begin_commit()
            os.link(candidate, output)
            published = True
            log.update(status='verified', output_sha256=digest(output), verification='approved immutable snapshot exact bytes, DOCX reopen, deterministic package checks and renderer')
            if log['output_sha256'] != plan['candidate_sha256']:
                raise GuardError('发布结果身份不符')
            (job/'audit.json').write_text(json.dumps(log,ensure_ascii=False,indent=2),encoding='utf8')
            return log
        except Exception as error:
            if published:
                log['status'] = 'uncertain'
            elif isinstance(error, PublicationCancelled):
                log['failure_code'] = 'publication_cancelled'
            elif isinstance(error, PublicationRejected):
                log['failure_code'] = error.reason
            (job/'audit.json').write_text(json.dumps(log,ensure_ascii=False,indent=2),encoding='utf8')
            raise
        finally:
            candidate.unlink(missing_ok=True)
