from datetime import datetime, timezone
from hashlib import sha256
from typing import Literal

from pydantic import BaseModel, ConfigDict, Field


class Record(BaseModel):
    model_config = ConfigDict(extra='forbid', strict=True, frozen=True)


class LiteratureItem(Record):
    record_type: Literal['literature_item'] = 'literature_item'
    id: str
    title: str
    authors: tuple[str, ...] = ()
    year: int | None = None
    doi: str | None = None
    identifiers: tuple[str, ...] = ()
    source_url: str | None = None
    retrieved_at: str
    access: Literal['metadata', 'abstract', 'public_fulltext_available', 'user_local'] = 'metadata'
    conflicts: tuple[str, ...] = ()
    retraction_status: Literal['unknown', 'reported_retracted'] = 'unknown'


class SourceDocument(Record):
    record_type: Literal['source_document'] = 'source_document'
    id: str
    source_name: str
    source_sha256: str = Field(pattern=r'^[0-9a-f]{64}$')
    parser_version: str
    scope_id: str
    retrieved_at: str
    access: Literal['user_local', 'public_fulltext'] = 'user_local'
    source_url: str | None = None
    literature_id: str | None = None


class EvidenceFragment(Record):
    record_type: Literal['evidence_fragment'] = 'evidence_fragment'
    id: str = Field(pattern=r'^ev:[0-9a-f]{64}$')
    document_id: str
    source_sha256: str = Field(pattern=r'^[0-9a-f]{64}$')
    location: str
    heading_path: tuple[str, ...] = ()
    page: int | None = None
    bbox: tuple[float, float, float, float] | None = None
    original_text: str = Field(max_length=8000)
    context: str = Field(default='', max_length=8000)
    summary: str | None = None
    extraction_quality: Literal['native_text', 'pdf_content_order', 'ocr_unverified'] = 'native_text'
    chunk_version: str = 'structure-v1'
    embedding_model: str | None = None
    embedding_version: str | None = None
    embedding_dimension: int | None = None
    index_version: str | None = None
    writable_locator_id: str | None = None


class ClaimSupport(Record):
    record_type: Literal['claim_support'] = 'claim_support'
    claim: str = Field(min_length=1, max_length=2000)
    evidence_ids: tuple[str, ...]
    relation: Literal['supports', 'contradicts', 'insufficient']
    applicability: str = Field(min_length=1, max_length=2000)
    limitation: str = Field(min_length=1, max_length=2000)
    assessment: Literal['user_reviewed', 'unverified'] = 'unverified'


def identifier(*parts):
    return sha256('\0'.join(parts).encode('utf-8')).hexdigest()


def acquired_at():
    return datetime.now(timezone.utc).isoformat()
