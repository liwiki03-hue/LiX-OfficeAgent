import hashlib
import hmac
import json
from pathlib import Path
import re
import secrets
from typing import Literal

from pydantic import Field

from .professional_edits import ConstrainedDocx, LocalChange
from .provenance import ClaimSupport, EvidenceFragment, SourceDocument
from .writes import GuardError, Workspace

STYLE_SHA256 = 'b4c7619fc16c45a31e4cc3271eab94ffe83192d3b4c7fc729470a3b459448de3'
STYLE_VERSION = 'ieee@6ba2e9c6aa6e0dab87d3a20b833ed8ff6500f698/citeproc-py@0.11.1'
_SEAL_KEY = secrets.token_bytes(32)


def _encoded(value):
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(',', ':')).encode('utf-8')


def seal(value):
    return hmac.new(_SEAL_KEY, _encoded(value), hashlib.sha256).hexdigest()


def format_reference(metadata):
    from citeproc import CitationStylesStyle, CitationStylesBibliography, Citation, CitationItem, formatter
    from citeproc.source.json import CiteProcJSON
    from importlib.metadata import version
    style = Path(__file__).parent / 'assets/csl/ieee.csl'
    if hashlib.sha256(style.read_bytes()).hexdigest() != STYLE_SHA256 or version('citeproc-py') != '0.11.1':
        raise GuardError('professional_citation_version')
    bibliography = CitationStylesBibliography(CitationStylesStyle(str(style), validate=False), CiteProcJSON([metadata]), formatter.plain)
    citation = Citation([CitationItem(metadata['id'])])
    bibliography.register(citation)
    inline = str(bibliography.cite(citation, lambda item: None))
    entries = [str(item) for item in bibliography.bibliography()]
    if inline != '[1]' or len(entries) != 1 or len(entries[0]) > 6000 or not entries[0].startswith(inline):
        raise GuardError('professional_citation_invalid')
    return inline, inline + ' ' + entries[0][len(inline):].lstrip()


class GroundedDocx(ConstrainedDocx):
    kind: Literal['grounded_docx'] = 'grounded_docx'
    changes: list[LocalChange] = Field(min_length=2, max_length=2)
    proof: dict
    proof_seal: str = Field(pattern=r'^[0-9a-f]{64}$')
    support: ClaimSupport
    style_version: Literal[STYLE_VERSION] = STYLE_VERSION


def validate_source(proof, proof_seal, target_path):
    from .rag_sources import SourceExtractor
    from .rag_evidence import _dir
    from .provenance import identifier
    if not hmac.compare_digest(seal(proof), proof_seal):
        raise GuardError('professional_evidence_invalid')
    document = SourceDocument.model_validate(proof['document'])
    fragments = [EvidenceFragment.model_validate_json(json.dumps(row)) for row in proof['fragments']]
    root = Path(target_path).parent
    if document.scope_id != identifier(str(root), str(_dir(root))):
        raise GuardError('professional_evidence_changed')
    try:
        if proof.get('archive'):
            from .literature_archive import LiteratureArchive
            from .rag_evidence import _read
            path = LiteratureArchive(Workspace(root), 'pdf').check_pdf_receipt(proof['archive'])
            _, stamp = _read(path)
            fingerprint = stamp[-1]
        else:
            name = proof['filename']
            if not isinstance(name, str) or Path(name).name != name or '/' in name or '\\' in name or ':' in name or not name.lower().endswith('.pdf'):
                raise GuardError('professional_evidence_invalid')
            _, _, fingerprint = SourceExtractor(Workspace(root))._snapshot(name)
    except Exception:
        raise GuardError('professional_evidence_changed') from None
    if fingerprint != document.source_sha256:
        raise GuardError('professional_evidence_changed')
    recorded_parser = proof.get('coverage', {}).get('parser_identity')
    if recorded_parser:
        from .pdf_reader import parser_identity
        if parser_identity() != recorded_parser:
            raise GuardError('professional_parser_changed')
    if any(row.source_sha256 != fingerprint or row.document_id != document.id or
                            row.id != 'ev:' + identifier(fingerprint, row.location, row.original_text) for row in fragments):
        raise GuardError('professional_evidence_invalid')
    return fragments


def validate_proof(request):
    fragments = validate_source(request.proof, request.proof_seal, request.path)
    if request.proof.get('ocr') and request.proof['ocr'].get('reviewed') is not True:
        raise GuardError('professional_ocr_review_required')
    support = request.support
    if (support.assessment != 'user_reviewed' or support.relation == 'insufficient' or
            not all(value.strip() for value in (support.claim, support.applicability, support.limitation)) or
            not 1 <= len(support.evidence_ids) <= 3 or
            len(set(support.evidence_ids)) != len(support.evidence_ids) or
            any(key not in {row.id for row in fragments} for key in support.evidence_ids)):
        raise GuardError('professional_support_required')
    if not fragments:
        raise GuardError('professional_evidence_invalid')
    return [row for row in fragments if row.id in support.evidence_ids]


def transform(parts, request):
    from . import professional_edits as local
    from .rag import SECRET
    fragments = validate_proof(request)
    _, entries = local._inventory(parts, request.source_sha256)
    from lxml import etree as ET
    numbered = re.compile(r'\[\s*\d+(?:\s*[,;–-]\s*\d+)*\s*\]')
    if any(numbered.search(local._text(node)) for name, data in parts.items() if name.startswith('word/') and name.endswith('.xml')
           for node in ET.fromstring(data).iter(local.W + 'p')):
        raise GuardError('professional_existing_numbered_citations')
    changes = [local.LocalChange.model_validate(row) if type(row) is dict else row for row in request.changes]
    if changes[0].locator_id == changes[1].locator_id:
        raise GuardError('professional_overlapping_changes')
    targets = {row['locator_id']: row for row in entries}
    body, reference = (targets.get(change.locator_id) for change in changes)
    if not body or not reference or any(not re.fullmatch(r'body:\d+/paragraph', row['location']) or row['style'] for row in (body, reference)):
        raise GuardError('professional_grounded_target')
    if entries.index(body) >= entries.index(reference):
        raise GuardError('professional_grounded_target')
    inline, bibliography = format_reference(request.proof['csl'])
    claim = request.support.claim
    if changes[0].new_text != claim + ' ' + inline or changes[1].new_text != bibliography:
        raise GuardError('professional_citation_invalid')
    if any(SECRET.search(value) for value in (claim, request.support.applicability, request.support.limitation)):
        raise GuardError('professional_sensitive_text')
    if local.CITATION.search(claim):
        raise GuardError('professional_citation_invalid')
    for pattern in (local.NUMBER, local.FORMULA):
        old = pattern.findall(body['text'])
        new = pattern.findall(claim)
        cursor = iter(new)
        if any(not any(value == item for value in cursor) for item in old):
            raise GuardError('professional_protected_numbers' if pattern is local.NUMBER else 'professional_protected_formulas')
        quotes = ' '.join(row.original_text for row in fragments)
        if any(value not in old and value not in pattern.findall(quotes) for value in new):
            raise GuardError('professional_unsupported_fact')
        if pattern.search(reference['text']):
            raise GuardError('professional_grounded_target')
    if local.CITATION.search(body['text']) or local.CITATION.search(reference['text']):
        raise GuardError('professional_grounded_target')
    base = ConstrainedDocx(path=request.path, source_sha256=request.source_sha256, changes=[row.model_copy(update={'evidence_ids': []}) for row in changes],
                           preserve_numbers=False, preserve_citations=False, preserve_formulas=False)
    changed, before, after = local.transform(parts, base)
    for item in after:
        item.update(evidence_ids=list(request.support.evidence_ids), claim_support=request.support.model_dump(mode='json'),
                    citation_style=request.style_version, source_document=request.proof['document'],
                    evidence_fragments=[row.model_dump(mode='json') for row in fragments], csl=request.proof['csl'],
                    semantic_verification='user_reviewed_not_machine_proven')
    return changed, before, after
