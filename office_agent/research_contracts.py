from typing import Annotated, Literal

from pydantic import Field

from .writes import Strict


class EvidenceCard(Strict):
    claim: str = Field(min_length=1, max_length=2000)
    evidence_ids: list[str] = Field(max_length=3)
    relation: str = Field(pattern='^(supports|contradicts|insufficient)$')
    applicability: str = Field(max_length=2000)
    limitation: str = Field(max_length=2000)
    conflicts: list[str] = Field(default_factory=list, max_length=8)


class Suggestion(Strict):
    locator_id: str = Field(pattern=r'^loc:[0-9a-f]{64}$')
    new_text: str = Field(min_length=1, max_length=2000)
    reason: str = Field(min_length=1, max_length=400)
    evidence_ids: list[str] = Field(max_length=3)
    risks: list[str] = Field(max_length=8)
    needs_evidence: bool


class Suggestions(Strict):
    cards: list[EvidenceCard] = Field(max_length=6)
    suggestions: list[Suggestion] = Field(max_length=12)
    missing_evidence: list[str] = Field(max_length=12)
    stop_reason: str = Field(min_length=1, max_length=200)


class ResearchPlan(Strict):
    queries: list[Annotated[str, Field(min_length=1, max_length=200)]] = Field(min_length=1, max_length=3)
    read_ids: list[str] = Field(max_length=6)


class VerifiedClaimSuggestion(Strict):
    claim: str = Field(min_length=1, max_length=1000)
    evidence_ids: list[str] = Field(max_length=3)
    verdict: Literal['supports', 'contradicts', 'insufficient', 'condition_mismatch', 'needs_manual_judgment']
    applicability: str = Field(min_length=1, max_length=500)
    limitation: str = Field(min_length=1, max_length=500)
    recommendation: str = Field(min_length=1, max_length=500)
    comparison: str = Field(min_length=1, max_length=500)


class ClaimVerification(Strict):
    claims: list[VerifiedClaimSuggestion] = Field(min_length=1, max_length=3)
    suggestions: list[Suggestion] = Field(max_length=1)
    missing_evidence: list[str] = Field(max_length=8)
    stop_reason: str = Field(min_length=1, max_length=200)
