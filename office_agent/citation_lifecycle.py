import hashlib
import hmac
import json
from pathlib import Path
import re
import time
from typing import Literal

from lxml import etree as ET
from pydantic import Field

from .grounded_edits import STYLE_SHA256, STYLE_VERSION, seal, validate_proof
from .professional_edits import ConstrainedDocx, LocalChange, FORMULA
from .writes import GuardError

NUMBERED = re.compile(r'\[(\d{1,3}(?:\s*(?:,|;|[-–])\s*\d{1,3})*)\]')
UNUSED = '参考文献占位（未使用）'
VERSION = 'citation-lifecycle-v1'


def numbers(value):
    result = []
    for item in re.split(r'\s*[,;]\s*', value):
        limits = re.split(r'\s*[-–]\s*', item)
        if len(limits) > 2 or any(not part.isdecimal() for part in limits):
            raise GuardError('professional_citation_syntax')
        first, last = int(limits[0]), int(limits[-1])
        if not 1 <= first <= last <= 999 or last - first > 20:
            raise GuardError('professional_citation_syntax')
        result.extend(range(first, last + 1))
    if len(result) > 24:
        raise GuardError('professional_capacity')
    return list(dict.fromkeys(result))


def scan(text):
    visible = FORMULA.sub(lambda match: 'x' * len(match.group()), text)
    if any(NUMBERED.fullmatch(match.group()) is None for match in re.finditer(r'\[[\d\s,;–-]+\]', visible)):
        raise GuardError('professional_citation_syntax')
    return [(match, numbers(match.group(1))) for match in NUMBERED.finditer(visible)]


def format_document(metadata, clusters):
    from importlib.metadata import version
    from citeproc import CitationStylesStyle, CitationStylesBibliography, Citation, CitationItem, formatter
    from citeproc.source.json import CiteProcJSON
    style = Path(__file__).parent / 'assets/csl/ieee.csl'
    if hashlib.sha256(style.read_bytes()).hexdigest() != STYLE_SHA256 or version('citeproc-py') != '0.11.1':
        raise GuardError('professional_citation_version')
    bibliography = CitationStylesBibliography(CitationStylesStyle(str(style), validate=False), CiteProcJSON(metadata), formatter.plain)
    citations = [Citation([CitationItem(key) for key in cluster]) for cluster in clusters]
    for citation in citations:
        bibliography.register(citation)
    bibliography.sort()
    order = list(dict.fromkeys(key for cluster in clusters for key in cluster))
    if bibliography.keys != order:
        raise GuardError('professional_citation_invalid')
    inlines = [str(bibliography.cite(citation, lambda item: None)) for citation in citations]
    entries = [str(item) for item in bibliography.bibliography()]
    if len(entries) != len(order) or len(order) > 8:
        raise GuardError('professional_capacity')
    for cluster, rendered in zip(clusters, inlines):
        if [n for _, ns in scan(rendered) for n in ns] != sorted({order.index(key) + 1 for key in cluster}):
            raise GuardError('professional_citation_invalid')
    for index, value in enumerate(entries, 1):
        label = f'[{index}]'
        if not value.startswith(label) or len(value) > 6000:
            raise GuardError('professional_citation_invalid')
        entries[index - 1] = label + ' ' + value[len(label):].lstrip()
    return inlines, entries, order


class CitationDocx(ConstrainedDocx):
    kind: Literal['citation_docx'] = 'citation_docx'
    intent: dict
    intent_seal: str = Field(pattern=r'^[0-9a-f]{64}$')
    style_version: Literal[STYLE_VERSION] = STYLE_VERSION


def _guard_claim(before, claim, fragments):
    from . import professional_edits as local
    plain = before
    for match, _ in reversed(scan(before)):
        plain = plain[:match.start()] + plain[match.end():]
    if local.CITATION.findall(plain) != local.CITATION.findall(claim):
        raise GuardError('professional_protected_citations')
    quotes = ' '.join(row.original_text for row in fragments)
    for pattern, code in ((local.NUMBER, 'numbers'), (local.FORMULA, 'formulas')):
        old, new = pattern.findall(plain), pattern.findall(claim)
        cursor = iter(new)
        if any(not any(item == value for item in cursor) for value in old):
            raise GuardError('professional_protected_' + code)
        if any(value not in old and value not in pattern.findall(quotes) for value in new):
            raise GuardError('professional_unsupported_fact')


def reconcile(parts, request):
    from . import professional_edits as local
    from .rag import SECRET
    intent = request.intent
    if not hmac.compare_digest(seal(intent), request.intent_seal):
        raise GuardError('professional_evidence_invalid')
    if time.monotonic() >= intent['deadline']:
        raise GuardError('professional_plan_expired')
    if intent.get('parent_audit'):
        from .rag_evidence import _read
        from .writes import Workspace
        descriptor = intent['parent_audit']
        if not re.fullmatch('[a-f0-9]{32}', descriptor['receipt_id']):
            raise GuardError('professional_citation_receipt')
        try:
            root = Path(request.path).parent
            _, stamp = _read(Workspace(root).file(str(root / '.office_agent' / descriptor['receipt_id'] / 'audit.json')))
            if json.loads(json.dumps(stamp)) != descriptor['stamp']:
                raise ValueError()
        except Exception:
            raise GuardError('professional_citation_receipt') from None
    _, rows = local._inventory(parts, request.source_sha256)
    by_location = {row['location']: row for row in rows}
    allowed = intent['authorized']
    slots = intent['slots']
    bindings = intent['bindings']
    catalog = intent['catalog']
    if (not 1 <= len(allowed) <= 12 or len(set(allowed)) != len(allowed) or not 1 <= len(slots) <= 8 or
            len(set(slots)) != len(slots) or not set(slots) <= set(allowed) or not 1 <= len(catalog) <= 16):
        raise GuardError('professional_citation_scope')
    for location in allowed:
        row = by_location.get(location)
        if not row or not re.fullmatch(r'body:\d+/paragraph', location) or not row['writable'] or row['style']:
            raise GuardError('professional_grounded_target')
    for name, data in parts.items():
        if name.startswith('word/') and name.endswith('.xml'):
            for node in ET.fromstring(data).iter(local.W + 'p'):
                if scan(local._text(node)) and (name != 'word/document.xml' or node.getparent().tag != local.W + 'body'):
                    raise GuardError('professional_citation_outside_body')
    for location, row in by_location.items():
        if re.match(r'^\[\d+\]', row['text']) and re.search(r'10\.\d{4,9}/', row['text']) and location not in slots:
            raise GuardError('professional_citation_scope')
    bib_numbers = set()
    for location in slots:
        text = by_location[location]['text']
        match = re.match(r'^\[(\d{1,3})\]', text)
        if match:
            if match.group(1) in bib_numbers:
                raise GuardError('professional_bibliography_identity')
            bib_numbers.add(match.group(1))
            key = bindings.get(match.group(1))
            tokens = re.findall(r'(?<![\w.])10\.\d{4,9}/[^\s<>"{}\\]+', text.casefold())
            expected = catalog[key]['DOI'].casefold() if key in catalog else ''
            if not key or key not in catalog or not any(token in {expected, expected + '.', expected + ',', expected + ';'} for token in tokens):
                raise GuardError('professional_bibliography_identity')
        elif any(pattern.search(text) for pattern in (local.NUMBER, local.FORMULA, local.CITATION)):
            raise GuardError('professional_grounded_target')
    claims = {item['location']: item for item in intent['claims']}
    if len(claims) != len(intent['claims']) or len(claims) > 4 or set(claims) & set(slots) or not set(claims) <= set(allowed):
        raise GuardError('professional_citation_scope')
    removals = {item['location']: item['numbers'] for item in intent['removals']}
    if len(removals) != len(intent['removals']) or set(removals) & (set(slots) | set(claims)) or not set(removals) <= set(allowed):
        raise GuardError('professional_citation_scope')
    templates, clusters, supports = {}, [], []
    for row in rows:
        location, text = row['location'], row['text']
        if location in slots:
            continue
        if any(str(n) not in bindings or str(n) not in bib_numbers for _, ns in scan(text) for n in ns):
            raise GuardError('professional_citation_unmapped')
        if location in claims:
            claim = claims[location]
            fragments, keys = [], []
            for attribution in claim['attributions']:
                from .provenance import ClaimSupport
                support = ClaimSupport.model_validate_json(json.dumps(attribution['support']))
                if any(SECRET.search(value) for value in (support.claim, support.applicability, support.limitation)):
                    raise GuardError('professional_sensitive_text')
                if support.claim != claim['text']:
                    raise GuardError('professional_support_required')
                proof = attribution['proof']
                check = type('ProofRequest', (), {'proof': proof, 'proof_seal': attribution['seal'], 'path': request.path, 'support': support})()
                fragments.extend(validate_proof(check))
                key = 'doi:' + proof['csl']['DOI'].casefold()
                if key not in catalog or {k: v for k, v in proof['csl'].items() if k not in {'id', 'URL'}} != {k: v for k, v in catalog[key].items() if k not in {'id', 'URL'}}:
                    raise GuardError('professional_duplicate_metadata')
                keys.append(key)
            if not 1 <= len(keys) <= 3 or len(set(keys)) != len(keys) or scan(claim['text']):
                raise GuardError('professional_support_required')
            _guard_claim(text, claim['text'], fragments)
            templates[location] = [claim['text'] + ' ', len(clusters)]
            clusters.append(keys)
            supports.append(claim)
        else:
            pieces, cursor = [], 0
            present = {n for _, ns in scan(text) for n in ns}
            remove = removals.get(location, [])
            if not set(remove) <= present or len(set(remove)) != len(remove):
                raise GuardError('professional_citation_unmapped')
            for match, ns in scan(text):
                if any(str(n) not in bindings for n in ns):
                    raise GuardError('professional_citation_unmapped')
                keys = list(dict.fromkeys(bindings[str(n)] for n in ns if n not in remove))
                if any(key not in catalog for key in keys):
                    raise GuardError('professional_citation_unmapped')
                pieces.append(text[cursor:match.start()])
                if keys:
                    pieces.append(len(clusters))
                    clusters.append(keys)
                cursor = match.end()
            pieces.append(text[cursor:])
            templates[location] = pieces
    if not clusters:
        inlines, bibliography, order = [], [], []
    else:
        inlines, bibliography, order = format_document([dict(value, id=key) for key, value in catalog.items()], clusters)
    if len(bibliography) > len(slots):
        raise GuardError('professional_bibliography_capacity')
    ordered_slots = sorted(slots, key=lambda location: next(i for i, row in enumerate(rows) if row['location'] == location))
    if any(i > rows.index(by_location[ordered_slots[0]]) for i, row in enumerate(rows) if row['location'] not in slots and scan(row['text'])) or any(rows.index(by_location[location]) >= rows.index(by_location[ordered_slots[0]]) for location in claims):
        raise GuardError('professional_grounded_target')
    final = {location: ''.join(inlines[piece] if type(piece) is int else piece for piece in pieces) for location, pieces in templates.items()}
    final.update({location: bibliography[index] if index < len(bibliography) else UNUSED for index, location in enumerate(ordered_slots)})
    changes = []
    for row in rows:
        location = row['location']
        if final.get(location, row['text']) == row['text']:
            continue
        if location not in allowed:
            raise GuardError('professional_citation_scope')
        changes.append(LocalChange(locator_id=row['locator_id'], expected_text_sha256=row['text_sha256'], new_text=final[location],
                                   reason='显式授权范围内重建引文编号、书目及已审阅主张'))
    if not changes:
        raise GuardError('professional_no_changes')
    manifest = {'version': VERSION, 'style': STYLE_VERSION, 'catalog': {key: catalog[key] for key in order},
                'number_to_key': {str(index): key for index, key in enumerate(order, 1)}, 'slots': ordered_slots,
                'previous_number_to_key': bindings,
                'number_changes': {number: order.index(key) + 1 if key in order else None for number, key in bindings.items()},
                'citation_locations': [location for location, text in final.items() if location not in slots and scan(text)],
                'removals': intent['removals'], 'source_sha256': request.source_sha256,
                'removed_keys': [key for key in catalog if key not in order], 'claims': supports,
                'parent_receipt': intent.get('parent_receipt'), 'archive_dependencies': intent.get('archive_dependencies', []),
                'historical_support': 'not_reassessed; retained in parent receipt'}
    return changes, manifest


def transform(parts, request):
    from . import professional_edits as local
    changes, manifest = reconcile(parts, request)
    if [item.model_dump() for item in changes] != [item.model_dump() for item in request.changes]:
        raise GuardError('professional_citation_invalid')
    base = ConstrainedDocx(path=request.path, source_sha256=request.source_sha256, changes=changes,
                           preserve_numbers=False, preserve_formulas=False, preserve_citations=False)
    changed, before, after = local.transform(parts, base)
    after[0]['citation_manifest'] = manifest
    after[0]['claim_support'] = 'user_reviewed_not_machine_proven'
    return changed, before, after
