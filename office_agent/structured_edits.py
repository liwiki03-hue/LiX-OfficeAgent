from copy import deepcopy
from datetime import datetime, timezone
import hmac
import json
import re
from typing import Literal

from lxml import etree as ET
from pydantic import Field

from .writes import Strict, GuardError
from .grounded_edits import seal, validate_proof
from .professional_edits import W, XML, LocalChange, ConstrainedDocx, _inventory, _canonical, _text, _rewrite
from .citation_lifecycle import scan, format_document, STYLE_VERSION
from .rag import SECRET

REL = '{http://schemas.openxmlformats.org/package/2006/relationships}'
CT = '{http://schemas.openxmlformats.org/package/2006/content-types}'
RECEIPT_VERSION = 'structured-citations-v1'


class ParagraphInsertion(Strict):
    anchor: str = Field(pattern=r'^loc:[a-f0-9]{64}$')
    side: Literal['before', 'after']
    text: str = Field(min_length=1, max_length=2000)
    style: str | None = Field(default=None, max_length=100)
    attributions: list[dict] = Field(default_factory=list, max_length=3)


class ReviewComment(Strict):
    anchor: str = Field(pattern=r'^loc:[a-f0-9]{64}$')
    start: int = Field(ge=0, le=8000)
    end: int = Field(ge=1, le=8000)
    quote: str = Field(min_length=1, max_length=2000)
    text: str = Field(min_length=1, max_length=2000)


class StructuredDocx(Strict):
    kind: Literal['structured_docx'] = 'structured_docx'
    path: str
    source_sha256: str = Field(pattern=r'^[a-f0-9]{64}$')
    insertions: list[ParagraphInsertion] = Field(default_factory=list, max_length=4)
    comments: list[ReviewComment] = Field(default_factory=list, max_length=8)
    replacements: list[LocalChange] = Field(default_factory=list, max_length=12)
    intent: dict
    intent_seal: str = Field(pattern=r'^[a-f0-9]{64}$')


def anchor_record(row, entries, version):
    index = entries.index(row)
    return {'source_sha256': version, 'locator_id': row['locator_id'], 'location': row['location'],
            'node': _canonical(row['node']).hex(),
            'previous': _canonical(entries[index - 1]['node']).hex() if index else None,
            'next': _canonical(entries[index + 1]['node']).hex() if index + 1 < len(entries) else None}


def _plain(value):
    if not value.strip() or SECRET.search(value) or any(ord(c) < 32 for c in value):
        raise GuardError('professional_sensitive_or_control_text')


def validate_parent(request):
    from .pdf_table import validate_proofs
    validate_proofs(request)
    intent = request.intent
    if intent.get('parent_audit'):
        from pathlib import Path
        from .rag_evidence import _read
        from .writes import Workspace
        try:
            parent = intent['parent_audit']
            if not re.fullmatch('[a-f0-9]{32}', parent['receipt_id']):
                raise ValueError()
            space = Workspace(Path(request.path).parent)
            data, stamp = _read(space.file(str(space.root / '.office_agent' / parent['receipt_id'] / 'audit.json')))
            if json.loads(json.dumps(stamp)) != parent['stamp'] or json.loads(data)['status'] != 'verified':
                raise ValueError()
        except Exception:
            raise GuardError('professional_citation_receipt') from None


def transform(parts, request):
    intent = request.intent
    if not hmac.compare_digest(seal(intent), request.intent_seal) or intent.get('source_sha256') != request.source_sha256:
        raise GuardError('professional_structure_authority')
    validate_parent(request)
    if intent['insertions'] != [r.model_dump() for r in request.insertions] or intent['comments'] != [r.model_dump() for r in request.comments]:
        raise GuardError('professional_structure_authority')
    if intent.get('replacements', []) != [r.model_dump() for r in request.replacements]:
        raise GuardError('professional_structure_authority')
    if not request.insertions and not request.comments and not request.replacements:
        raise GuardError('professional_no_changes')
    if request.comments and any('comment' in name.casefold() for name in parts):
        raise GuardError('professional_existing_comments')
    root, entries = _inventory(parts, request.source_sha256)
    targets = {r['locator_id']: r for r in entries}
    frozen_anchors = {r['locator_id']: anchor_record(r, entries, request.source_sha256) for r in entries}
    allowed_styles = {None}
    if 'word/styles.xml' in parts:
        styles = ET.fromstring(parts['word/styles.xml'])
        by_style = {s.get(W + 'styleId'): s for s in styles.findall(W + 'style') if s.get(W + 'type') == 'paragraph'}
        def ordinary_style(key, seen=()):
            if key in seen or key not in by_style:
                return False
            node = by_style[key]
            if any(node.find(W + 'pPr/' + W + tag) is not None for tag in ('numPr', 'sectPr', 'outlineLvl')):
                return False
            parent = node.find(W + 'basedOn')
            return parent is None or ordinary_style(parent.get(W + 'val'), (*seen, key))
        allowed_styles.update(key for key in by_style if ordinary_style(key))
        if any(s.get(W + 'default') == '1' and not ordinary_style(key) for key, s in by_style.items()):
            allowed_styles.discard(None)
    def target(key):
        row = targets.get(key)
        if not row or not row['writable'] or not re.fullmatch(r'body:\d+/paragraph', row['location']):
            raise GuardError('professional_structure_anchor')
        if row['node'].find(W + 'pPr/' + W + 'numPr') is not None or row['node'].find(W + 'pPr/' + W + 'sectPr') is not None:
            raise GuardError('professional_structure_anchor')
        if frozen_anchors[key] != intent['anchors'].get(key):
            raise GuardError('professional_stale_locator')
        return row
    before, after, insert_nodes, seen = [], [], [], set()
    original_nodes = [(r['node'], _canonical(r['node'])) for r in entries]
    replacement_ids = {r.locator_id for r in request.replacements}
    if replacement_ids & ({r.anchor for r in request.comments} | set(intent['bibliography'])):
        raise GuardError('professional_overlapping_changes')
    if request.replacements:
        from .professional_edits import transform as local_transform
        updated, replaced_before, replaced_after = local_transform(parts, ConstrainedDocx(path=request.path, source_sha256=request.source_sha256, changes=request.replacements))
        _, new_entries = _inventory(updated, request.source_sha256)
        by_location = {r['location']: r for r in new_entries}
        for row in entries:
            if row['locator_id'] in replacement_ids:
                new = by_location[row['location']]['node']
                for child in list(row['node']):
                    row['node'].remove(child)
                row['node'].extend(deepcopy(list(new)))
        before.extend(dict(r, operation='replace') for r in replaced_before)
        after.extend(dict(r, operation='replace') for r in replaced_after)
    tails = {}
    citation_keys = {}
    insertion_receipts = {}
    for insertion in request.insertions:
        row = target(insertion.anchor)
        _plain(insertion.text)
        if scan(insertion.text):
            raise GuardError('professional_citation_syntax')
        signature = (insertion.anchor, insertion.side, insertion.text)
        if signature in seen:
            raise GuardError('professional_overlapping_changes')
        seen.add(signature)
        style = insertion.style if insertion.style is not None else row['style']
        if style not in allowed_styles:
            raise GuardError('professional_structure_style')
        node = ET.Element(W + 'p')
        if style:
            ET.SubElement(ET.SubElement(node, W + 'pPr'), W + 'pStyle', {W + 'val': style})
        run = ET.SubElement(node, W + 'r')
        first = row['node'].find(W + 'r/' + W + 'rPr')
        if first is not None:
            run.append(deepcopy(first))
        ET.SubElement(run, W + 't', {XML + 'space': 'preserve'}).text = insertion.text
        anchor = row['node']
        if insertion.side == 'before':
            anchor.addprevious(node)
        else:
            tails.get(insertion.anchor, anchor).addnext(node)
            tails[insertion.anchor] = node
        keys, fragments = [], []
        for attribution in insertion.attributions:
            from .provenance import ClaimSupport
            support = ClaimSupport.model_validate_json(json.dumps(attribution['support']))
            check = type('ProofRequest', (), dict(proof=attribution['proof'], proof_seal=attribution['seal'], path=request.path, support=support))()
            fragments.extend(validate_proof(check))
            if support.claim != insertion.text:
                raise GuardError('professional_support_required')
            keys.append('doi:' + attribution['proof']['csl']['DOI'].casefold())
        if keys:
            from .citation_lifecycle import _guard_claim
            _guard_claim('', insertion.text, fragments)
            citation_keys[node] = list(dict.fromkeys(keys))
        insert_nodes.append(node)
        before.append({'operation': 'insert', 'anchor': row['location'], 'side': insertion.side, 'text': None})
        after.append({'operation': 'insert', 'anchor': row['location'], 'side': insertion.side, 'text': insertion.text,
                      'mode': 'clean_paragraph_insertion', 'support': 'user_reviewed_not_machine_proven' if keys else 'not_assessed'})
        after[-1]['attributions'] = [{'source_document': a['proof']['document'], 'evidence_fragments': a['proof']['fragments'],
                                     'claim_support': a['support'], 'csl': a['proof']['csl'], 'coverage': a['proof'].get('coverage'),
                                     'archive_save_id': a['proof'].get('archive', {}).get('save_id')}
                                    for a in insertion.attributions]
        insertion_receipts[node] = after[-1]
    catalog, bindings = intent['catalog'], intent['bindings']
    slots = [target(key) for key in intent['bibliography']]
    if len({r['locator_id'] for r in slots}) != len(slots):
        raise GuardError('professional_citation_scope')
    slots.sort(key=lambda row: entries.index(row))
    if len(catalog) > 8:
        raise GuardError('professional_capacity')
    if citation_keys:
        slot_nodes = {r['node'] for r in slots}
        clusters, templates = [], {}
        body = root.find(W + 'body')
        for row in entries:
            if row['node'] in slot_nodes:
                continue
            if scan(row['text']) and row['node'].getparent() is not body:
                raise GuardError('professional_citation_scope')
        for node in body:
            if node.tag != W + 'p' or node in slot_nodes:
                continue
            if node in citation_keys:
                templates[node] = [_text(node) + ' ', len(clusters)]
                clusters.append(citation_keys[node])
                continue
            text, pieces, cursor = _text(node), [], 0
            for match, ns in scan(text):
                if any(str(n) not in bindings for n in ns):
                    raise GuardError('professional_citation_unmapped')
                pieces.extend([text[cursor:match.start()], len(clusters)])
                clusters.append(list(dict.fromkeys(bindings[str(n)] for n in ns)))
                cursor = match.end()
            pieces.append(text[cursor:])
            templates[node] = pieces
        if any(key not in catalog for cluster in clusters for key in cluster):
            raise GuardError('professional_citation_unmapped')
        for number, key in bindings.items():
            if not re.fullmatch('[1-9][0-9]{0,2}', number) or key not in catalog:
                raise GuardError('professional_citation_unmapped')
        if len(slots) != len(bindings):
            raise GuardError('professional_bibliography_identity')
        for row in slots:
            label = re.match(r'^\[(\d+)\]\s', row['text'])
            if not label or label[1] not in bindings or catalog[bindings[label[1]]]['DOI'].casefold() not in row['text'].casefold():
                raise GuardError('professional_bibliography_identity')
        inline, bibliography, order = format_document([dict(v, id=k) for k, v in catalog.items()], clusters)
        if len(bibliography) < len(slots):
            raise GuardError('professional_citation_scope')
        bib_anchor = target(intent['bibliography_anchor'])['node']
        if any(node in templates and (scan(_text(node)) or node in citation_keys) for node in list(body)[list(body).index(bib_anchor) + 1:] if node not in slot_nodes):
            raise GuardError('professional_grounded_target')
        for node, pieces in templates.items():
            text = ''.join(inline[p] if type(p) is int else p for p in pieces)
            if text != _text(node):
                if node not in insert_nodes:
                    row = next(r for r in entries if r['node'] is node)
                    if row['locator_id'] not in intent['authorized'] or not row['writable']:
                        raise GuardError('professional_citation_scope')
                    before.append({'operation': 'renumber', 'location': row['location'], 'text': row['text']})
                    after.append({'operation': 'renumber', 'location': row['location'], 'text': text})
                _rewrite(node, _text(node), text)
                if node in insertion_receipts:
                    insertion_receipts[node]['text'] = text
        bibliography_nodes = []
        for index, text in enumerate(bibliography):
            if index < len(slots):
                row = slots[index]
                if row['text'] != text:
                    if row['locator_id'] not in intent['authorized']:
                        raise GuardError('professional_citation_scope')
                    before.append({'operation': 'bibliography_update', 'location': row['location'], 'text': row['text']})
                    after.append({'operation': 'bibliography_update', 'location': row['location'], 'text': text})
                    _rewrite(row['node'], row['text'], text)
                bib_anchor = row['node']
                bibliography_nodes.append(bib_anchor)
            else:
                node = ET.Element(W + 'p')
                style = slots[-1]['style'] if slots else None
                if style not in allowed_styles:
                    raise GuardError('professional_structure_style')
                if style:
                    ET.SubElement(ET.SubElement(node, W + 'pPr'), W + 'pStyle', {W + 'val': style})
                ET.SubElement(ET.SubElement(node, W + 'r'), W + 't').text = text
                bib_anchor.addnext(node)
                bib_anchor = node
                bibliography_nodes.append(node)
                insert_nodes.append(node)
                before.append({'operation': 'bibliography_insert', 'text': None})
                after.append({'operation': 'bibliography_insert', 'text': text})
        after.append({'operation': 'citation_map', 'number_to_key': {str(i): key for i, key in enumerate(order, 1)},
                      'previous_number_to_key': bindings, 'catalog': {key: catalog[key] for key in order}})
    for note in intent.get('review_notes', []):
        if type(note) is not str or len(note) > 300:
            raise GuardError('professional_invalid_request')
        _plain(note)
    if intent.get('review_notes'):
        after.append({'operation': 'review_notes', 'ignored_reasons': intent['review_notes']})
    for review in intent.get('pdf_numeric_reviews', []):
        after.append({'operation': 'pdf_numeric_review', 'review': review})
    changed = dict(parts)
    if request.comments:
        if any('comment' in name.casefold() for name in parts) or any(n.tag in {W + 'commentRangeStart', W + 'commentReference'} for n in root.iter()):
            raise GuardError('professional_existing_comments')
        comments = ET.Element(W + 'comments', nsmap={'w': W[1:-1]})
        occupied = {}
        from .tracked_edits import _slices
        for index, item in enumerate(request.comments):
            row = target(item.anchor)
            _plain(item.text)
            if _text(row['node']) != row['text'] or not 0 <= item.start < item.end <= len(row['text']) or row['text'][item.start:item.end] != item.quote:
                raise GuardError('professional_comment_range')
            if item.anchor in occupied:
                raise GuardError('professional_overlapping_changes')
            occupied[item.anchor] = True
            node, runs = row['node'], list(row['node'].findall(W + 'r'))
            for run in runs:
                node.remove(run)
            for run in _slices(runs, 0, item.start):
                node.append(run)
            ET.SubElement(node, W + 'commentRangeStart', {W + 'id': str(index)})
            for run in _slices(runs, item.start, item.end):
                node.append(run)
            ET.SubElement(node, W + 'commentRangeEnd', {W + 'id': str(index)})
            ET.SubElement(ET.SubElement(node, W + 'r'), W + 'commentReference', {W + 'id': str(index)})
            for run in _slices(runs, item.end, len(row['text'])):
                node.append(run)
            comment = ET.SubElement(comments, W + 'comment', {W + 'id': str(index), W + 'author': 'Office Agent', W + 'date': intent['date']})
            ET.SubElement(ET.SubElement(ET.SubElement(comment, W + 'p'), W + 'r'), W + 't').text = item.text
            before.append({'operation': 'comment', 'location': row['location'], 'quote': item.quote})
            after.append({'operation': 'comment', 'location': row['location'], 'quote': item.quote, 'text': item.text, 'author': 'Office Agent', 'date': intent['date']})
        rel_name = 'word/_rels/document.xml.rels'
        rels = ET.fromstring(parts[rel_name])
        if any('comments' in (r.get('Type', '') + r.get('Target', '')) for r in rels):
            raise GuardError('professional_existing_comments')
        ids = {r.get('Id') for r in rels}
        rid = next('rId' + str(i) for i in range(1, len(ids) + 2) if 'rId' + str(i) not in ids)
        ET.SubElement(rels, REL + 'Relationship', Id=rid, Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments', Target='comments.xml')
        types = ET.fromstring(parts['[Content_Types].xml'])
        ET.SubElement(types, CT + 'Override', PartName='/word/comments.xml', ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml')
        for name, node in [('word/comments.xml', comments), (rel_name, rels), ('[Content_Types].xml', types)]:
            changed[name] = ET.tostring(node, xml_declaration=True, encoding='UTF-8', standalone=True)
    changed['word/document.xml'] = ET.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    if citation_keys or intent.get('prior_citations'):
        _, output_rows = _inventory(changed, request.source_sha256)
        if not citation_keys:
            prior = intent['prior_citations']
            catalog = prior['catalog']
            order = [prior['number_to_key'][str(i)] for i in range(1, len(prior['number_to_key']) + 1)]
            bibliography_nodes = [next(r['node'] for r in entries if r['location'] == location) for location in prior['slots']]
            body = root.find(W + 'body')
        slots = ['body:' + str(list(body).index(node)) + '/paragraph' for node in bibliography_nodes]
        manifest = {'version': RECEIPT_VERSION, 'style': STYLE_VERSION,
                    'catalog': {key: catalog[key] for key in order},
                    'number_to_key': {str(i): key for i, key in enumerate(order, 1)},
                    'slots': slots, 'parent_receipt': intent.get('parent_receipt'),
                    'archive_dependencies': intent.get('archive_dependencies', []),
                    'historical_support': 'not_reassessed; retained in parent receipt',
                    'targets': [{'location': r['location'], 'text': r['text']} for r in output_rows if scan(r['text'])]}
        after.append({'operation': 'citation_receipt', 'citation_manifest': manifest})
        restore_citation_receipt({'items': [after[-1]]}, output_rows)
    affected = set(intent['authorized']) | {c.anchor for c in request.comments} | replacement_ids
    for row, (node, original) in zip(entries, original_nodes):
        if row['locator_id'] not in affected and _canonical(node) != original:
            raise GuardError('professional_non_target_structure_changed')
    if sum(len(json.dumps(v, ensure_ascii=False)) for v in after) > 32000:
        raise GuardError('professional_capacity')
    return changed, before, after


def restore_citation_receipt(after, rows):
    manifests = [r['citation_manifest'] for r in after['items'] if 'citation_manifest' in r]
    if len(manifests) != 1:
        raise ValueError()
    manifest = manifests[0]
    if manifest['version'] != RECEIPT_VERSION or manifest['style'] != STYLE_VERSION:
        raise ValueError()
    targets = [{'location': r['location'], 'text': r['text']} for r in rows if scan(r['text'])]
    if manifest['targets'] != targets or not 1 <= len(manifest['slots']) <= 8:
        raise ValueError()
    order = [manifest['number_to_key'][str(i)] for i in range(1, len(manifest['number_to_key']) + 1)]
    if len(set(order)) != len(order) or set(order) != set(manifest['catalog']) or len(manifest['slots']) != len(order):
        raise ValueError()
    _, bibliography, actual_order = format_document([dict(v, id=k) for k, v in manifest['catalog'].items()], [[k] for k in order])
    by_location = {r['location']: r for r in rows}
    if actual_order != order or len(set(manifest['slots'])) != len(order):
        raise ValueError()
    for location, text in zip(manifest['slots'], bibliography):
        row = by_location[location]
        if not re.fullmatch(r'body:\d+/paragraph', location) or row['text'] != text:
            raise ValueError()
    if any(str(n) not in manifest['number_to_key'] for row in rows for _, ns in scan(row['text']) for n in ns):
        raise ValueError()
    return manifest


def verify(before, after, request):
    expected, _, _ = transform(before, request)
    allowed = {'word/document.xml'}
    if request.comments:
        allowed |= {'word/comments.xml', 'word/_rels/document.xml.rels', '[Content_Types].xml'}
    if set(after) != set(expected) or any(before.get(k) != after.get(k) for k in before.keys() | after.keys() if k not in allowed):
        raise GuardError('professional_non_target_part_changed')
    for name in allowed:
        if name in expected and _canonical(ET.fromstring(expected[name])) != _canonical(ET.fromstring(after[name])):
            raise GuardError('professional_non_target_structure_changed')
