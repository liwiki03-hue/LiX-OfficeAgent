import hmac
import json
from pathlib import Path
from typing import Literal

from pydantic import Field

from .grounded_edits import seal, validate_source
from .professional_edits import ConstrainedDocx
from .rag import SECRET
from .writes import GuardError


class AssistedDocx(ConstrainedDocx):
    kind: Literal['assisted_docx'] = 'assisted_docx'
    review: dict
    review_seal: str = Field(pattern=r'^[0-9a-f]{64}$')


def validate_review(request):
    review = request.review
    if not hmac.compare_digest(seal(review), request.review_seal) or review['source_sha256'] != request.source_sha256:
        raise GuardError('professional_assist_review_invalid')
    if review['mode'] != 'polish' or len(review['edits']) != len(request.changes):
        raise GuardError('professional_assist_review_required')
    for edit, change in zip(review['edits'], request.changes):
        if edit['reviewed'] is not True or edit['locator_id'] != change.locator_id or edit['new_text'] != change.new_text:
            raise GuardError('professional_assist_review_required')
    if SECRET.search(json.dumps(review, ensure_ascii=False)):
        raise GuardError('professional_sensitive_text')
    for item in review['proofs']:
        validate_source(item['proof'], item['seal'], request.path)
        if item['proof'].get('ocr') and item['proof']['ocr'].get('reviewed') is not True:
            raise GuardError('professional_ocr_review_required')
    if review.get('rag_receipt'):
        from .rag_evidence import _read
        from .rag_sources import SourceExtractor
        from .writes import Workspace
        root = Path(request.path).parent
        descriptor = review['rag_receipt']
        import re
        if not re.fullmatch('[a-f0-9]{32}', descriptor['save_id']):
            raise GuardError('professional_evidence_invalid')
        _, stamp = _read(root / ('.office-rag-evidence-' + descriptor['save_id']) / 'evidence.json')
        if json.loads(json.dumps(stamp)) != descriptor['stamp']:
            raise GuardError('professional_evidence_changed')
        extractor = SourceExtractor(Workspace(root))
        for evidence in review['selected_evidence'].values():
            if evidence['capture_id'] is None:
                _, _, fingerprint = extractor._snapshot(evidence['title'])
                if fingerprint != evidence['source_sha256']:
                    raise GuardError('professional_evidence_changed')


def transform(parts, request):
    from . import professional_edits as local
    validate_review(request)
    base = ConstrainedDocx(**request.model_dump(exclude={'kind', 'review', 'review_seal'}))
    changed, before, after = local.transform(parts, base)
    after[0]['assisted_review'] = request.review
    after[0]['semantic_verification'] = 'human_reviewed_edit_not_machine_proven'
    return changed, before, after
